# 02 — Refactor ChatProvider to be App-Agnostic

## Goal

The `ChatProvider` in `chat-context.tsx` is the main glue between the shared UI and the app-specific logic. Currently it hardcodes:

1. `EXCEL_TOOLS` — the tool list
2. `buildSystemPrompt()` — the system prompt
3. `getWorkbookMetadata()` — app context injected into user messages
4. `<wb_context>` tag — the XML wrapper for metadata
5. `navigateTo()` — follow mode (Excel-specific)
6. `DirtyRange` parsing — Excel-specific dirty tracking
7. `stripEnrichment()` pattern — strips `<wb_context>` from display

These need to become configurable via props/config so each app can inject its own.

## Approach: AppConfig Interface

Define an `AppConfig` that each app passes to `ChatProvider`:

```typescript
// packages/shared/src/chat/types.ts

import type { AgentTool } from "@mariozechner/pi-agent-core";
import type { SkillMeta } from "../skills";

export interface AppConfig {
  /** Display name shown in the UI */
  appName: string;

  /** The tools this app provides to the agent */
  tools: AgentTool[];

  /**
   * Build the system prompt.
   * Receives installed skills so the app can include them.
   */
  buildSystemPrompt: (skills: SkillMeta[]) => string;

  /**
   * Fetch app-specific context to inject into user messages.
   * For Excel: workbook metadata (sheets, active cell, etc.)
   * For PPT: presentation state (slide count, masters, dimensions, etc.)
   * Returns null if unavailable.
   */
  getAppContext?: () => Promise<{ tag: string; content: string } | null>;

  /**
   * Handle tool result side effects (e.g. Excel follow mode).
   * Called after each successful tool execution.
   * Optional — apps that don't need this can omit it.
   */
  onToolResult?: (toolCallId: string, resultText: string) => void;

  /**
   * App-specific settings to add to the provider config.
   * e.g. Excel has `followMode`, PPT might have something else.
   */
  extraConfigDefaults?: Record<string, unknown>;

  /**
   * Custom commands for the VFS bash shell.
   * e.g. Excel has csv-to-sheet, sheet-to-csv.
   */
  customCommands?: import("just-bash/browser").CustomCommand[];

  /**
   * Storage configuration.
   */
  storage: {
    dbName: string;
    dbVersion: number;
    settingsKey: string;
    storageKeyPrefix: string;
  };
}
```

## Changes to ChatProvider

### Constructor

```tsx
// Before:
export function ChatProvider({ children }: { children: ReactNode }) {

// After:
export function ChatProvider({ config, children }: { config: AppConfig; children: ReactNode }) {
```

### System Prompt

```typescript
// Before:
function buildSystemPrompt(skills: SkillMeta[]): string {
  return `You are an AI assistant integrated into Microsoft Excel...`;
}

// After: use config.buildSystemPrompt(skills)
const systemPrompt = config.buildSystemPrompt(skillsRef.current);
```

### Tools

```typescript
// Before:
tools: EXCEL_TOOLS,

// After:
tools: config.tools,
```

### App Context (metadata injection)

```typescript
// Before:
const metadata = await getWorkbookMetadata();
promptContent = `<wb_context>\n${JSON.stringify(metadata)}\n</wb_context>\n\n${content}`;

// After:
if (config.getAppContext) {
  const ctx = await config.getAppContext();
  if (ctx) {
    promptContent = `<${ctx.tag}>\n${ctx.content}\n</${ctx.tag}>\n\n${content}`;
  }
}
```

### Follow Mode / Tool Result Handling

```typescript
// Before: hardcoded dirty range parsing + navigateTo()
if (!event.isError && followModeRef.current) {
  const dirtyRanges = parseDirtyRanges(resultText);
  // ... navigateTo(first.sheetId, first.range)
}

// After:
if (!event.isError && config.onToolResult) {
  config.onToolResult(event.toolCallId, resultText);
}
```

### Strip Enrichment

```typescript
// Before: hardcodes <wb_context> pattern
text = text.replace(/^<wb_context>\n[\s\S]*?\n<\/wb_context>\n\n/, "");

// After: generic — strip any <tag>...</tag> block at the start of user messages
text = text.replace(/^<[a-z_]+>\n[\s\S]*?\n<\/[a-z_]+>\n\n/, "");
```

Or better, use the `tag` from `AppConfig.getAppContext` to build the regex.

### VFS Custom Commands

```typescript
// Before:
import { getCustomCommands } from "./custom-commands";
customCommands: getCustomCommands(),

// After: shared commands + app commands
import { getSharedCommands } from "@open-office/shared/vfs";
customCommands: [...getSharedCommands(), ...(config.customCommands ?? [])],
```

### Storage Init

```typescript
// Before: hardcoded in db.ts
openDB("OpenExcelDB_v3", 30, ...)

// After: init from config at ChatProvider mount
useEffect(() => {
  initStorage(config.storage);
  initProviderConfig(config.storage.storageKeyPrefix);
  initOAuth(config.storage.storageKeyPrefix);
}, []);
```

## What Stays App-Specific in ChatProvider

Almost nothing. After this refactor, `ChatProvider` becomes a generic component in shared that:
- Takes an `AppConfig`
- Manages agent lifecycle, streaming, sessions, file uploads, skills
- Renders the chat UI

## Changes to chat-interface.tsx

Minimal:
- `ChatInterface` accepts and passes through `AppConfig`
- The `followMode` toggle only renders if `config.onToolResult` is defined
- The stats bar, session dropdown, theme toggle — all stay as-is

```tsx
// Before:
export function ChatInterface() {
  return <ChatProvider><ChatContent /></ChatProvider>;
}

// After:
export function ChatInterface({ config }: { config: AppConfig }) {
  return <ChatProvider config={config}><ChatContent /></ChatProvider>;
}
```

## Changes to settings-panel.tsx

The settings panel is 99% generic. The only Excel-specific bit is:
- The `followMode` toggle (line ~700ish)

Make it conditional: only render if `config.extraConfigDefaults?.followMode !== undefined` or if the app signals it supports follow mode.

## Changes to message-list.tsx

Check if there are any Excel-specific citation patterns (`#cite:sheetId`). If so, the citation click handler needs to become app-specific (passed via config or context).

Currently `message-list.tsx` has a link click handler for `#cite:` URLs that calls `navigateTo()`. This should move to `AppConfig`:

```typescript
export interface AppConfig {
  // ... existing fields ...
  
  /** Handle citation link clicks. Return true if handled. */
  onCitationClick?: (hash: string) => boolean;
}
```

## Checklist

- [ ] Define `AppConfig` interface in `packages/shared/src/chat/types.ts`
- [ ] Refactor `ChatProvider` to accept `AppConfig`
- [ ] Replace hardcoded `EXCEL_TOOLS` with `config.tools`
- [ ] Replace hardcoded `buildSystemPrompt` with `config.buildSystemPrompt`
- [ ] Replace `getWorkbookMetadata` with `config.getAppContext`
- [ ] Replace follow-mode/dirty-range handling with `config.onToolResult`
- [ ] Make `stripEnrichment` generic
- [ ] Pass `AppConfig` through `ChatInterface`
- [ ] Conditionalize `followMode` toggle in settings
- [ ] Move citation click handler to `AppConfig.onCitationClick`
- [ ] Init storage/oauth/provider-config from `config.storage`
