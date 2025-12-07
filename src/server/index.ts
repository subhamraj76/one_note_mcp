#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { MsalClient } from '../auth/msal-client';
import { OneNoteClient, RateLimitError } from '../graph/onenote-client';
import { markdownToHtml, htmlToMarkdown } from '../utils/markdown';

// Get cache directory from environment variable (set by extension)
const CACHE_DIR = process.env.ONENOTE_MCP_CACHE_DIR || './.vscode';

// Initialize clients
let msalClient: MsalClient | null = null;
let oneNoteClient: OneNoteClient | null = null;

async function getOneNoteClient(): Promise<OneNoteClient> {
  if (!msalClient) {
    msalClient = new MsalClient(
      { cacheDir: CACHE_DIR },
      (message) => console.error(`[Warning] ${message}`)
    );
    await msalClient.initialize();
  }

  if (!oneNoteClient) {
    oneNoteClient = new OneNoteClient(() => msalClient!.getAccessToken());
  }

  return oneNoteClient;
}

function isRateLimitError(result: any): result is RateLimitError {
  return result && typeof result === 'object' && 'isRateLimited' in result && result.isRateLimited === true;
}

function formatRateLimitResponse(error: RateLimitError): string {
  return JSON.stringify({
    error: 'rate_limited',
    message: error.message,
    retryAfterSeconds: error.retryAfterSeconds
  }, null, 2);
}

// Create the MCP server
const server = new McpServer({
  name: 'onenote-mcp',
  version: '1.0.0'
});

// Tool 1: Search Notebooks
server.tool(
  'search_notebooks',
  'Search for OneNote notebooks by name. Returns matching notebook IDs and names for use with other tools.',
  {
    query: z.string().describe('Search query to match against notebook names (case-insensitive fuzzy match)')
  },
  async ({ query }) => {
    try {
      const client = await getOneNoteClient();
      const notebooks = await client.searchNotebooks(query);

      if (isRateLimitError(notebooks)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(notebooks) }],
          isError: true
        };
      }

      if (notebooks.length === 0) {
        return {
          content: [{ type: 'text', text: `No notebooks found matching "${query}"` }]
        };
      }

      const result = notebooks.map(nb => ({
        id: nb.id,
        name: nb.displayName,
        isDefault: nb.isDefault,
        lastModified: nb.lastModifiedDateTime
      }));

      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error searching notebooks: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 2: Get Notebook Sections
server.tool(
  'get_notebook_sections',
  'List all sections in a specific OneNote notebook. Use the notebook ID from search_notebooks.',
  {
    notebook_id: z.string().describe('The ID of the notebook to get sections from')
  },
  async ({ notebook_id }) => {
    try {
      const client = await getOneNoteClient();
      const sections = await client.getSections(notebook_id);

      if (isRateLimitError(sections)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(sections) }],
          isError: true
        };
      }

      if (sections.length === 0) {
        return {
          content: [{ type: 'text', text: 'No sections found in this notebook' }]
        };
      }

      const result = sections.map(s => ({
        id: s.id,
        name: s.displayName,
        lastModified: s.lastModifiedDateTime
      }));

      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error getting sections: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 3: Get Section Pages
server.tool(
  'get_section_pages',
  'List all pages in a specific OneNote section. Use the section ID from get_notebook_sections.',
  {
    section_id: z.string().describe('The ID of the section to get pages from')
  },
  async ({ section_id }) => {
    try {
      const client = await getOneNoteClient();
      const pages = await client.getPages(section_id);

      if (isRateLimitError(pages)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(pages) }],
          isError: true
        };
      }

      if (pages.length === 0) {
        return {
          content: [{ type: 'text', text: 'No pages found in this section' }]
        };
      }

      const result = pages.map(p => ({
        id: p.id,
        title: p.title,
        lastModified: p.lastModifiedDateTime
      }));

      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error getting pages: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 4: Read Page
server.tool(
  'read_page',
  'Read the content of a OneNote page. Returns the page content converted to Markdown for easy reading.',
  {
    page_id: z.string().describe('The ID of the page to read')
  },
  async ({ page_id }) => {
    try {
      const client = await getOneNoteClient();
      const htmlContent = await client.getPageContent(page_id);

      if (isRateLimitError(htmlContent)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(htmlContent) }],
          isError: true
        };
      }

      // Convert HTML to Markdown for better readability
      const markdown = htmlToMarkdown(htmlContent);

      return {
        content: [{ type: 'text', text: markdown }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error reading page: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 5: Search OneNote
server.tool(
  'search_onenote',
  'Search for text across all OneNote pages. Returns matching pages with their IDs and titles.',
  {
    query: z.string().describe('Search query to find in page content'),
    scope: z.string().optional().describe('Optional: Limit search to specific notebook or section ID')
  },
  async ({ query, scope }) => {
    try {
      const client = await getOneNoteClient();
      const results = await client.searchPages(query, scope);

      if (isRateLimitError(results)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(results) }],
          isError: true
        };
      }

      if (results.length === 0) {
        return {
          content: [{ type: 'text', text: `No pages found matching "${query}"` }]
        };
      }

      const formattedResults = results.map(r => ({
        id: r.id,
        title: r.title,
        preview: r.preview
      }));

      return {
        content: [{ type: 'text', text: JSON.stringify(formattedResults, null, 2) }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error searching: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 6: Create Page
server.tool(
  'create_page',
  'Create a new OneNote page in a section. Content should be in Markdown format - it will be converted to HTML automatically. Supports Mermaid diagrams.',
  {
    section_id: z.string().describe('The ID of the section to create the page in'),
    title: z.string().describe('Title for the new page'),
    content_markdown: z.string().describe('Page content in Markdown format. Mermaid diagrams (```mermaid blocks) are supported.')
  },
  async ({ section_id, title, content_markdown }) => {
    try {
      const client = await getOneNoteClient();
      
      // Convert Markdown to HTML
      const htmlContent = markdownToHtml(content_markdown);
      
      const page = await client.createPage(section_id, title, htmlContent);

      if (isRateLimitError(page)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(page) }],
          isError: true
        };
      }

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: `Page "${title}" created successfully`,
            pageId: page.id,
            createdAt: page.createdDateTime
          }, null, 2)
        }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error creating page: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Tool 7: Update Page (Append)
server.tool(
  'update_page',
  'Append content to an existing OneNote page. Content should be in Markdown format - it will be converted to HTML and appended to the page body.',
  {
    page_id: z.string().describe('The ID of the page to update'),
    content_markdown: z.string().describe('Content to append in Markdown format. Mermaid diagrams (```mermaid blocks) are supported.')
  },
  async ({ page_id, content_markdown }) => {
    try {
      const client = await getOneNoteClient();
      
      // Convert Markdown to HTML
      const htmlContent = markdownToHtml(content_markdown);
      
      const success = await client.appendToPage(page_id, htmlContent);

      if (isRateLimitError(success)) {
        return {
          content: [{ type: 'text', text: formatRateLimitResponse(success) }],
          isError: true
        };
      }

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Content appended to page successfully',
            pageId: page_id
          }, null, 2)
        }]
      };
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error updating page: ${error.message}` }],
        isError: true
      };
    }
  }
);

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('OneNote MCP Server started');
}

main().catch((error) => {
  console.error('Failed to start OneNote MCP Server:', error);
  process.exit(1);
});
