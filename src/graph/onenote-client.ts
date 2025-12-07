import { Client } from '@microsoft/microsoft-graph-client';

export interface Notebook {
  id: string;
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isDefault?: boolean;
}

export interface Section {
  id: string;
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  parentNotebookId?: string;
}

export interface Page {
  id: string;
  title: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  contentUrl?: string;
  parentSectionId?: string;
}

export interface SearchResult {
  id: string;
  title: string;
  preview?: string;
  parentSectionId?: string;
  parentNotebookId?: string;
}

export interface RateLimitError {
  isRateLimited: boolean;
  retryAfterSeconds?: number;
  message: string;
}

interface GraphError {
  statusCode?: number;
  code?: string;
  message?: string;
}

const MAX_RETRIES = 3;
const INITIAL_DELAY_MS = 1000;

async function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export class OneNoteClient {
  private client: Client;
  private accessTokenProvider: () => Promise<string>;

  constructor(accessTokenProvider: () => Promise<string>) {
    this.accessTokenProvider = accessTokenProvider;
    
    this.client = Client.init({
      authProvider: async (done) => {
        try {
          const token = await this.accessTokenProvider();
          done(null, token);
        } catch (error: any) {
          done(error, null);
        }
      }
    });
  }

  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    operationName: string
  ): Promise<T | RateLimitError> {
    let lastError: any;
    
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      try {
        return await operation();
      } catch (error: any) {
        lastError = error;
        
        const graphError = error as GraphError;
        
        // Check for rate limiting (429) or server errors (5xx)
        if (graphError.statusCode === 429 || 
            (graphError.statusCode && graphError.statusCode >= 500)) {
          
          // Get retry-after header if available
          const retryAfter = error.headers?.get?.('retry-after');
          const delayMs = retryAfter 
            ? parseInt(retryAfter, 10) * 1000 
            : INITIAL_DELAY_MS * Math.pow(2, attempt);
          
          if (attempt < MAX_RETRIES - 1) {
            await delay(delayMs);
            continue;
          }
          
          // Return rate limit info to agent after all retries exhausted
          return {
            isRateLimited: true,
            retryAfterSeconds: Math.ceil(delayMs / 1000),
            message: `Rate limited on ${operationName}. Please wait ${Math.ceil(delayMs / 1000)} seconds before retrying. Original error: ${graphError.message || 'Unknown error'}`
          };
        }
        
        // For other errors, don't retry
        throw error;
      }
    }
    
    throw lastError;
  }

  private isRateLimitError(result: any): result is RateLimitError {
    return result && typeof result === 'object' && 'isRateLimited' in result;
  }

  async listNotebooks(): Promise<Notebook[] | RateLimitError> {
    return this.executeWithRetry(async () => {
      const response = await this.client
        .api('/me/onenote/notebooks')
        .select('id,displayName,createdDateTime,lastModifiedDateTime,isDefault')
        .get();
      
      return response.value as Notebook[];
    }, 'listNotebooks');
  }

  async searchNotebooks(query: string): Promise<Notebook[] | RateLimitError> {
    const notebooks = await this.listNotebooks();
    
    if (this.isRateLimitError(notebooks)) {
      return notebooks;
    }
    
    const lowerQuery = query.toLowerCase();
    return notebooks.filter(nb => 
      nb.displayName.toLowerCase().includes(lowerQuery)
    );
  }

  async getSections(notebookId: string): Promise<Section[] | RateLimitError> {
    return this.executeWithRetry(async () => {
      const response = await this.client
        .api(`/me/onenote/notebooks/${notebookId}/sections`)
        .select('id,displayName,createdDateTime,lastModifiedDateTime')
        .get();
      
      return response.value.map((s: any) => ({
        ...s,
        parentNotebookId: notebookId
      })) as Section[];
    }, 'getSections');
  }

  async getPages(sectionId: string): Promise<Page[] | RateLimitError> {
    return this.executeWithRetry(async () => {
      const response = await this.client
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .select('id,title,createdDateTime,lastModifiedDateTime,contentUrl')
        .get();
      
      return response.value.map((p: any) => ({
        ...p,
        parentSectionId: sectionId
      })) as Page[];
    }, 'getPages');
  }

  async getPageContent(pageId: string): Promise<string | RateLimitError> {
    return this.executeWithRetry(async () => {
      // Use includeIDs=true for potential future updates
      const response = await this.client
        .api(`/me/onenote/pages/${pageId}/content`)
        .query({ includeIDs: 'true' })
        .get();
      
      // Response is HTML content
      if (typeof response === 'string') {
        return response;
      }
      
      // If it's a stream/buffer, convert to string
      if (Buffer.isBuffer(response)) {
        return response.toString('utf-8');
      }
      
      // If it's an ArrayBuffer or similar
      if (response instanceof ArrayBuffer) {
        return Buffer.from(response).toString('utf-8');
      }
      
      return String(response);
    }, 'getPageContent');
  }

  async searchPages(query: string, scope?: string): Promise<SearchResult[] | RateLimitError> {
    return this.executeWithRetry(async () => {
      // Build the search query
      let apiPath = '/me/onenote/pages';
      
      const response = await this.client
        .api(apiPath)
        .search(query)
        .select('id,title,parentSection')
        .top(25)
        .get();
      
      return response.value.map((p: any) => ({
        id: p.id,
        title: p.title,
        parentSectionId: p.parentSection?.id
      })) as SearchResult[];
    }, 'searchPages');
  }

  async createPage(
    sectionId: string, 
    title: string, 
    htmlContent: string
  ): Promise<Page | RateLimitError> {
    return this.executeWithRetry(async () => {
      const pageHtml = `
<!DOCTYPE html>
<html>
  <head>
    <title>${this.escapeHtml(title)}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
    ${htmlContent}
  </body>
</html>`;

      const response = await this.client
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header('Content-Type', 'text/html')
        .post(pageHtml);
      
      return {
        id: response.id,
        title: response.title,
        createdDateTime: response.createdDateTime,
        lastModifiedDateTime: response.lastModifiedDateTime,
        parentSectionId: sectionId
      } as Page;
    }, 'createPage');
  }

  async appendToPage(pageId: string, htmlContent: string): Promise<boolean | RateLimitError> {
    return this.executeWithRetry(async () => {
      // PATCH operation to append content to the page body
      const patchContent = [
        {
          target: 'body',
          action: 'append',
          content: htmlContent
        }
      ];

      await this.client
        .api(`/me/onenote/pages/${pageId}/content`)
        .header('Content-Type', 'application/json')
        .patch(patchContent);
      
      return true;
    }, 'appendToPage');
  }

  private escapeHtml(text: string): string {
    const htmlEntities: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    };
    return text.replace(/[&<>"']/g, char => htmlEntities[char] || char);
  }
}
