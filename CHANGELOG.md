# Changelog

## [Unreleased]

### Fixes

- **search_data pagination with offset > 0** — Requests with `offset > 0` could return zero matches even when matches exist, and `hasMore`/`nextOffset` could be incorrect. Extracted pagination logic into a pure `SearchPageCollector` with separate match counting and page collection.

### Chores

- Added `pnpm test` step to CI workflow.
- Removed redundant typecheck/lint from release workflow (already validated in CI).

## [0.2.1] - 2026-02-08

### Fixes

- **OAuth token refresh during agent loops** — Token was only refreshed once at the start of a message, so multi-turn tool-use conversations could fail mid-stream if the access token expired. Token refresh now happens before every LLM call inside `streamFn`, matching pi's `AuthStorage.getApiKey()` pattern.

## [0.2.0] - 2026-02-08

### Features

- **Virtual filesystem & bash shell** — In-memory VFS powered by `just-bash/browser`. The agent can now read/write files and execute sandboxed bash commands (pipes, redirections, loops) with output truncation.
- **File uploads & drag-and-drop** — Upload files via paperclip button or drag-and-drop onto chat. Files are written to `/home/user/uploads/` and persisted per session in IndexedDB.
- **Composable CLI commands** — `csv-to-sheet`, `sheet-to-csv`, `pdf-to-text`, `docx-to-text`, `xlsx-to-csv` bridge the VFS and Excel for data import/export.
- **OAuth authentication** — Anthropic (Claude Pro/Max) and OpenAI Codex (ChatGPT Plus/Pro) OAuth via PKCE flow with token refresh.
- **Custom endpoints** — Connect to any OpenAI-compatible API (Ollama, vLLM, LMStudio) or other supported API types with configurable base URL and API type.
- **Skills system** — Install agent skills (folders or single `SKILL.md` files with YAML frontmatter). Skills are persisted in IndexedDB, mounted into the VFS, and injected into the system prompt.

### Breaking Changes

- **Message storage migrated** — Sessions now store raw `AgentMessage[]` instead of derived `ChatMessage[]`. Old sessions will appear empty after upgrade.

### Improvements

- Context window usage in stats bar now shows actual context sent per turn (not cumulative totals).
- Scroll handler in message list switched from `addEventListener` to React `onScroll`.

### Chores

- Replaced Dexie with `idb` for IndexedDB access — Dexie's global `Promise` patching is incompatible with SES `lockdown()`, which froze `Promise` and broke all DB operations after `eval_officejs` was used.
- Removed dead scaffold files (`hero-list.tsx`, `text-insertion.tsx`, `header.tsx`).
- Removed old crypto shims (no longer needed with Vite polyfills).
- IndexedDB schema upgraded to v3 with `vfsFiles` and `skillFiles` tables.

## [0.1.10] - 2026-02-06

Initial release with AI chat interface, multi-provider LLM support (BYOK), Excel read/write tools, and CORS proxy configuration.
