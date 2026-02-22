# Changelog

## [Unreleased]

## [0.2.4] - 2026-02-22

### Features

- **Screenshot tool** — New `screenshot_range` tool captures Excel cell ranges as images and stores them in the VFS. New `pixart` CLI command renders pixel art from a simple text DSL directly into Excel cells.
- **Flexible set_cell_range** — `set_cell_range` now auto-pads rows with empty strings when row lengths don't match, removing the strict rectangular shape requirement.

### Fixes

- **Manifest validation** — Fixed invalid dev manifest GUID that prevented sideloading in Excel. Added manifest validation to `pnpm check` and CI.

### Chores

- Bumped `@mariozechner/pi-ai` and `@mariozechner/pi-agent-core` dependencies.

## [0.2.3] - 2026-02-15

### Features

- **Web search** — New `web-search` CLI command lets the agent search the web with pagination, region, and time filters. Supports multiple search providers: DuckDuckGo (free, no key), Brave, Serper, and Exa (API key required).
- **Web fetch** — New `web-fetch` CLI command fetches a URL and saves readable content to a file. HTML pages are converted to Markdown via Readability + Turndown. Binary files (PDF, DOCX, etc.) are downloaded raw. Supports basic fetch and Exa as providers.
- **Web tools settings** — New "Web Tools" section in the settings panel to configure search/fetch providers and manage API keys (Brave, Serper, Exa) with an advanced keys drawer.

### Improvements

- **Chat input redesign** — Input field now auto-resizes (up to 2 rows), with paperclip and send buttons moved inside the input border for a cleaner look.
- **Error boundary** — Added a top-level React error boundary that catches unhandled render errors and offers "Try again" / "Reload add-in" actions instead of a blank screen.

## [0.2.2] - 2026-02-13

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
