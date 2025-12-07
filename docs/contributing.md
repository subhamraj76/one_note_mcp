# Contributing Guide

A teaching-style guide for new contributors. No prior MCP, Graph, or VS Code extension experience required.

---

## 1. Who This Is For
- Engineers adding tools or fixing bugs.
- Docs contributors improving guides.
- QA / testers validating flows.

## 2. Prerequisites
- Node.js 20+
- Git
- VS Code 1.96+
- A Microsoft account with OneNote access

## 3. One-Time Setup
```bash
git clone https://github.com/rashmirrout/OneNoteMCP.git
cd OneNoteMCP
npm install
```

## 4. Daily Dev Loop (Fast Path)
1. Start watch build: `npm run watch`
2. Press `F5` in VS Code → launches Extension Development Host
3. In the new window, open Command Palette → run OneNote commands or let your AI agent call the MCP tools
4. Edit code → watch rebuilds → press `Ctrl+Shift+F5` to reload extension host

## 5. Branch & Commit Hygiene
- Create a feature branch: `git checkout -b feature/short-title`
- Keep commits small and focused; prefer present-tense imperative messages (`Add retry logging`).
- Run `npm run compile` (or `npm run watch`) before pushing.

## 6. Coding Standards
- Language: TypeScript (ES2022 target, Node 16 module).
- Lint: `npm run lint` (uses ESLint + @typescript-eslint).
- Keep functions small; handle errors with clear messages surfaced to agents.
- Prefer zod for validating tool inputs.
- Avoid logging PII; keep outputs user-friendly.

## 7. Testing & Verification
- Build once: `npm run compile`
- Watch: `npm run watch`
- Package (optional): `npx @vscode/vsce package`
- Manual test: use F5 Extension Development Host, then exercise tools via Command Palette or AI agent.
- If you add new Graph calls, consider mocking in future tests (MSW/nock) to avoid hitting Graph in CI.

## 8. Auth & Security Notes
- Auth is PKCE via localhost:3000; tokens cached under `.vscode/onenote-mcp-cache.json` (workspace) or global storage.
- OS-protected cache first (DPAPI/Keychain/libsecret); plaintext fallback with warning.
- **Commands available:**
  - `OneNote MCP: Check Auth Status` – shows ✅/⚠️ with account info
  - `OneNote MCP: Sign In` – clears cache and prepares re-auth
  - `OneNote MCP: Sign Out` – clears cache file completely
- To test auth flow, delete the cache file and use any tool to trigger login.

## 9. When Adding a New MCP Tool
1. Define parameters with zod in `src/server/index.ts`.
2. Implement handler using `OneNoteClient` or new Graph call.
3. Use `executeWithRetry` pattern; surface `RateLimitError` with retry hints.
4. Return JSON or Markdown that agents can render easily.
5. Rebuild and test via F5.

## 10. Documentation Updates
- Add architecture/design/auth changes to `docs/` (see existing examples).
- Link new docs from README if generally useful.

## 11. Pull Request Checklist
- [ ] Code compiles (`npm run compile`)
- [ ] Lint passes (`npm run lint`)
- [ ] Manual sanity test via F5 (at least one tool call)
- [ ] Docs updated if behavior changed
- [ ] No secrets committed; no PII in logs

## 12. Release (maintainers)
- Bump version in `package.json` if needed.
- `npm run package` then `npx @vscode/vsce package` to produce VSIX.
- Tag a release; CI workflow attaches VSIX to GitHub Release.
