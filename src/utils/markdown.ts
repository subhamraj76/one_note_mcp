import { marked } from 'marked';
import pako from 'pako';

/**
 * Converts Markdown content to OneNote-compatible HTML.
 * Handles special cases like Mermaid diagrams.
 */
export function markdownToHtml(markdown: string): string {
  // First, process Mermaid diagrams before passing to marked
  const processedMarkdown = processMermaidBlocks(markdown);
  
  // Convert markdown to HTML
  const html = marked.parse(processedMarkdown, {
    async: false,
    gfm: true,
    breaks: true
  }) as string;
  
  return html;
}

/**
 * Detects ```mermaid code blocks and converts them to image tags
 * using mermaid.ink service.
 */
function processMermaidBlocks(markdown: string): string {
  // Match mermaid code blocks
  const mermaidRegex = /```mermaid\n([\s\S]*?)```/g;
  
  return markdown.replace(mermaidRegex, (match, mermaidCode) => {
    const trimmedCode = mermaidCode.trim();
    const imageUrl = getMermaidImageUrl(trimmedCode);
    
    // Return an image tag that will render the diagram
    return `![Mermaid Diagram](${imageUrl})`;
  });
}

/**
 * Converts Mermaid diagram code to a mermaid.ink URL.
 * Uses pako compression for efficient URL encoding.
 */
function getMermaidImageUrl(mermaidCode: string): string {
  // Create the state object that mermaid.ink expects
  const state = {
    code: mermaidCode,
    mermaid: {
      theme: 'default'
    },
    autoSync: true,
    updateDiagram: true
  };
  
  // Compress using pako
  const jsonString = JSON.stringify(state);
  const compressed = pako.deflate(new TextEncoder().encode(jsonString), { level: 9 });
  
  // Convert to base64url (URL-safe base64)
  const base64 = btoa(String.fromCharCode(...compressed))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
  
  return `https://mermaid.ink/img/pako:${base64}`;
}

/**
 * Converts HTML to a simplified Markdown representation.
 * Used for reading page content.
 */
export function htmlToMarkdown(html: string): string {
  // Simple HTML to Markdown conversion
  // This is a basic implementation - could use turndown for production
  
  let markdown = html;
  
  // Remove HTML doctype, head, etc.
  markdown = markdown.replace(/<!DOCTYPE[^>]*>/gi, '');
  markdown = markdown.replace(/<head[\s\S]*?<\/head>/gi, '');
  markdown = markdown.replace(/<html[^>]*>/gi, '');
  markdown = markdown.replace(/<\/html>/gi, '');
  markdown = markdown.replace(/<body[^>]*>/gi, '');
  markdown = markdown.replace(/<\/body>/gi, '');
  
  // Convert common HTML elements to Markdown
  // Headers
  markdown = markdown.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '# $1\n\n');
  markdown = markdown.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '## $1\n\n');
  markdown = markdown.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '### $1\n\n');
  markdown = markdown.replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, '#### $1\n\n');
  markdown = markdown.replace(/<h5[^>]*>([\s\S]*?)<\/h5>/gi, '##### $1\n\n');
  markdown = markdown.replace(/<h6[^>]*>([\s\S]*?)<\/h6>/gi, '###### $1\n\n');
  
  // Bold and italic
  markdown = markdown.replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, '**$1**');
  markdown = markdown.replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, '**$1**');
  markdown = markdown.replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, '*$1*');
  markdown = markdown.replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, '*$1*');
  
  // Links
  markdown = markdown.replace(/<a[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
  
  // Images
  markdown = markdown.replace(/<img[^>]*src="([^"]*)"[^>]*alt="([^"]*)"[^>]*\/?>/gi, '![$2]($1)');
  markdown = markdown.replace(/<img[^>]*src="([^"]*)"[^>]*\/?>/gi, '![]($1)');
  
  // Lists
  markdown = markdown.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '- $1\n');
  markdown = markdown.replace(/<\/?ul[^>]*>/gi, '\n');
  markdown = markdown.replace(/<\/?ol[^>]*>/gi, '\n');
  
  // Paragraphs and line breaks
  markdown = markdown.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, '$1\n\n');
  markdown = markdown.replace(/<br\s*\/?>/gi, '\n');
  markdown = markdown.replace(/<div[^>]*>([\s\S]*?)<\/div>/gi, '$1\n');
  
  // Code blocks
  markdown = markdown.replace(/<pre[^>]*><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, '```\n$1\n```\n');
  markdown = markdown.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, '`$1`');
  
  // Clean up remaining tags
  markdown = markdown.replace(/<[^>]+>/g, '');
  
  // Decode HTML entities
  markdown = markdown
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
  
  // Clean up extra whitespace
  markdown = markdown.replace(/\n{3,}/g, '\n\n');
  markdown = markdown.trim();
  
  return markdown;
}

/**
 * Wraps HTML content in OneNote-compliant page structure.
 */
export function wrapInOneNoteHtml(title: string, bodyHtml: string): string {
  const escapedTitle = title
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
  
  return `<!DOCTYPE html>
<html>
  <head>
    <title>${escapedTitle}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
    ${bodyHtml}
  </body>
</html>`;
}
