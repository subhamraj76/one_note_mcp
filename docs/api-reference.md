# OneNote MCP – API Reference (Tools & Contracts)

> Detailed, teaching-friendly reference for every MCP tool, inputs/outputs, errors, and auth scopes.

---

## 1. Transport & Host
- **Transport**: MCP stdio via VS Code MCP provider.
- **Server binary**: `dist/server.js` (spawned by extension with `node`).
- **Env passed**: `ONENOTE_MCP_CACHE_DIR` (cache location).

---

## 2. Authentication
- OAuth2 + PKCE, loopback redirect `http://localhost:3000`.
- Scopes: `Notes.Read`, `Notes.ReadWrite`, `offline_access`, `openid`, `profile`.
- Tokens cached per workspace `.vscode/onenote-mcp-cache.json` (or global storage if no workspace) with OS-protected storage when available.

### Checking Auth Status
- **VS Code Command:** `OneNote MCP: Check Auth Status` → shows ✅ or ⚠️ with account email
- **Output Panel:** Select "OneNote MCP" dropdown for auth logs
- **Cache File:** Check if `onenote-mcp-cache.json` exists with content

### Forcing Re-Authentication
- **VS Code Command:** `OneNote MCP: Sign In` → clears cache, next tool call triggers login
- **Manual:** Delete cache file → next tool call triggers login

---

## 3. Common Response Shapes
- **Success (data)**: `content: [{ type: 'text', text: <Markdown or JSON string> }]`
- **Error**: `content: [{ type: 'text', text: <message> }], isError: true`
- **RateLimitError** (special): JSON string of `{ error: "rate_limited", message, retryAfterSeconds }`

---

## 4. Tools

### 4.1 `search_notebooks`
- **Description**: Search for OneNote notebooks by name.
- **Input**:
  - `query` (string) – search text (case-insensitive).
- **Output**: JSON array of `{ id, name, isDefault, lastModified }`.
- **Notes**: Uses client-side fuzzy match over `/me/onenote/notebooks`.

### 4.2 `get_notebook_sections`
- **Description**: List sections in a notebook.
- **Input**:
  - `notebook_id` (string)
- **Output**: JSON array of `{ id, name, lastModified }`.

### 4.3 `get_section_pages`
- **Description**: List pages in a section.
- **Input**:
  - `section_id` (string)
- **Output**: JSON array of `{ id, title, lastModified }`.

### 4.4 `read_page`
- **Description**: Read OneNote page content, returned as Markdown.
- **Input**:
  - `page_id` (string)
- **Output**: Markdown string of page content.
- **Notes**: HTML → Markdown conversion is basic; complex layouts may render simply.

### 4.5 `search_onenote`
- **Description**: Search across pages.
- **Input**:
  - `query` (string)
  - `scope` (optional string) – limit to a notebook or section ID.
- **Output**: JSON array of `{ id, title, preview }` (top 25).

### 4.6 `create_page`
- **Description**: Create a new page from Markdown.
- **Input**:
  - `section_id` (string)
  - `title` (string)
  - `content_markdown` (string) – supports ```mermaid blocks.
- **Output**: JSON `{ success, message, pageId, createdAt }`.
- **Notes**: Markdown → HTML with Mermaid converted to mermaid.ink images.

### 4.7 `update_page`
- **Description**: Append Markdown to an existing page.
- **Input**:
  - `page_id` (string)
  - `content_markdown` (string) – supports ```mermaid blocks.
- **Output**: JSON `{ success, message, pageId }`.

---

## 5. Error Handling & Throttling
- Retries: up to 3 attempts on 429/5xx with backoff (1s, 2s, 4s).
- If retries exhausted: returns `RateLimitError` object serialized to JSON with `retryAfterSeconds` when available.
- For other errors: returned as `isError: true` with message text.

---

## 6. Usage Examples (Pseudocode)

### Example: Read a page
```
call tool: read_page
params: { "page_id": "<PAGE_ID>" }
```
Response: Markdown string of the page content.

### Example: Create a page with Mermaid
```
call tool: create_page
params: {
  "section_id": "<SECTION_ID>",
  "title": "Sprint Plan",
  "content_markdown": """
# Sprint Plan

```mermaid
gantt
  dateFormat  YYYY-MM-DD
  title Sprint
  section Work
  Feature A :a1, 2025-01-01, 5d
  Feature B :after a1, 4d
```
"""
}
```

### Example: Handle rate limit
- If response contains JSON with `error: "rate_limited"`, wait `retryAfterSeconds` before retrying.

---

## 7. Operational Notes
- Cache dir is provided by the extension via env var; avoid hardcoding paths.
- `mermaid.ink` URLs are used for diagrams; no local Mermaid runtime needed.
- Outputs are strings (Markdown or JSON text). Clients should parse JSON text when needed.

---

## 8. Future Additions (Planned)
- More tools: move pages, copy sections, attachments upload.
- Richer search filters and semantic search.
- Higher-fidelity HTML→Markdown via `turndown`.
- Configurable redirect port for auth.
