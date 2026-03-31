# OneNote MCP - Complete Documentation

## Introduction
OneNote MCP is a Model Context Protocol (MCP) server packaged as a VS Code extension that enables AI agents (like GitHub Copilot and Claude) to interact with Microsoft OneNote. It allows reading, searching, and writing OneNote content via Microsoft Graph API.

**Key Features:**
- Zero-configuration authentication using Microsoft's public client ID
- Secure token caching with OS-native encryption
- Markdown/HTML conversion with Mermaid diagram support
- Rate limit handling with exponential backoff
- 7 MCP tools for OneNote operations

## Architecture
The system consists of:
- **VS Code Extension**: Hosts the MCP server definition provider
- **MCP Server**: Node.js stdio process exposing OneNote tools
- **Authentication Layer**: MSAL with PKCE for secure token management
- **Graph Client**: Wrapper for Microsoft Graph OneNote API calls
- **Markdown Adapter**: Converts between Markdown and OneNote HTML

### High-Level Flow
1. VS Code extension registers MCP server
2. AI agent calls MCP tools over stdio
3. Server authenticates via MSAL if needed
4. Graph API calls fetch/modify OneNote data
5. Responses converted to agent-friendly format

## Components

### VS Code Extension (`src/extension.ts`)
- Registers MCP server definition provider
- Provides commands: Check Auth Status, Sign In, Sign Out
- Launches MCP server process with environment variables

### MCP Server (`src/server/index.ts`)
- Uses `@modelcontextprotocol/sdk` with StdioServerTransport
- Implements 7 tools with Zod validation
- Handles rate limits and error responses

### Authentication (`src/auth/msal-client.ts`)
- MSAL Node with PKCE flow
- Public client ID: `14d82eec-204b-4c2f-b7e8-296a70dab67e`
- Loopback server on port 3000
- Token caching with OS protection (DPAPI/Keychain/libsecret)

### Graph Client (`src/graph/onenote-client.ts`)
- Wraps Microsoft Graph OneNote endpoints
- Exponential backoff retry (3 attempts: 1s, 2s, 4s)
- Detects rate limits and returns structured errors

### Markdown Utils (`src/utils/markdown.ts`)
- Markdown to HTML conversion using `marked`
- Mermaid diagram rendering via `mermaid.ink`
- Basic HTML to Markdown conversion

## MCP Tools

1. **search_notebooks**: Fuzzy search notebook names
   - Input: `query` (string)
   - Output: Array of `{id, name, isDefault, lastModified}`

2. **get_notebook_sections**: List sections in a notebook
   - Input: `notebook_id` (string)
   - Output: Array of `{id, name, lastModified}`

3. **get_section_pages**: List pages in a section
   - Input: `section_id` (string)
   - Output: Array of `{id, title, lastModified}`

4. **read_page**: Read page content as Markdown
   - Input: `page_id` (string)
   - Output: Markdown string

5. **search_onenote**: Global search across pages
   - Input: `query` (string), optional `scope`
   - Output: Array of `{id, title, preview}`

6. **create_page**: Create new page from Markdown
   - Input: `section_id`, `title`, `content_markdown`
   - Output: `{success, message, pageId, createdAt}`

7. **update_page**: Append Markdown to existing page
   - Input: `page_id`, `content_markdown`
   - Output: `{success, message, pageId}`

## Authentication Flow
- Uses OAuth2 + PKCE with Microsoft Graph scopes
- Access token: ~1 hour lifetime
- Refresh token: ~90 days (rolling)
- Cache locations:
  - Windows: DPAPI encrypted
  - macOS: Keychain
  - Linux: libsecret or plaintext with warning
- Login triggers on first use or token expiry

## Installation
1. Install VS Code extension from marketplace or VSIX
2. For development: clone repo, `npm install`, `npm run compile`
3. Configure MCP in VS Code or Claude Desktop

## Configuration

### VS Code (Extension)
- Automatic registration via MCP provider API
- Cache: `.vscode/onenote-mcp-cache.json`

### VS Code (Local Development)
- Create `.vscode/mcp.json`:
```json
{
  "servers": {
    "onenote-mcp": {
      "type": "stdio",
      "command": "node",
      "args": ["/path/to/dist/server.js"],
      "env": {"ONENOTE_MCP_CACHE_DIR": "${workspaceFolder}/.vscode"}
    }
  }
}
```

### Claude Desktop
- Edit `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "onenote-mcp": {
      "command": "node",
      "args": ["/path/to/dist/server.js"],
      "env": {"ONENOTE_MCP_CACHE_DIR": "/path/to/cache"}
    }
  }
}
```

## Usage
- In VS Code Copilot: `@workspace Find notebook "Work Notes"`
- In Claude: Use tools directly after configuration
- Authentication happens automatically on first tool call

## Code Structure
```
src/
├── extension.ts          # VS Code extension entry
├── server/
│   └── index.ts          # MCP server with tools
├── auth/
│   └── msal-client.ts    # Authentication client
├── graph/
│   └── onenote-client.ts # Graph API wrapper
└── utils/
    ├── index.ts
    └── markdown.ts       # Markdown/HTML conversion
```

## Dependencies
- `@modelcontextprotocol/sdk`: MCP implementation
- `@azure/msal-node`: Authentication
- `@microsoft/microsoft-graph-client`: Graph API
- `marked`: Markdown parsing
- `pako`: Compression for Mermaid
- `zod`: Schema validation

## Build Process
- `npm run compile`: Webpack build to `dist/`
- `npm run watch`: Watch mode for development
- `npm run package`: Production build with source maps
- `npm run lint`: ESLint checks

## Contributing
1. Prerequisites: Node.js 18+, Git, VS Code
2. Setup: `npm install`, `npm run compile`
3. Development: `npm run watch`, F5 debug
4. Testing: Manual testing with Copilot/Claude
5. PR: Follow coding standards, update docs

## Security
- No custom Azure app registration required
- Tokens cached with OS protection
- PKCE prevents auth code interception
- Rate limits handled gracefully
- No sensitive data logged

## Error Handling
- Rate limits: Retry with backoff, surface to agent
- Auth errors: Trigger re-authentication
- Graph errors: User-friendly messages
- Validation: Zod schemas for all inputs

## Extensibility
- Add tools in `server/index.ts` with Zod schemas
- Extend Graph client for new OneNote operations
- Modify Markdown converter for new formats
- Update webpack for additional bundles

## Troubleshooting
- Auth issues: Clear cache, check firewall for localhost:3000
- Tool errors: Check rate limits, token validity
- Build issues: Ensure Node.js version, clean node_modules
- MCP not found: Reload VS Code, check mcp.json path

## Future Improvements
- Support for OneNote images/attachments
- Batch operations for multiple pages
- Custom notebook templates
- Integration with other Microsoft services
- Enhanced Mermaid diagram support
