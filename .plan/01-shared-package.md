# 01 — Extract Shared Package

## Goal

Create `packages/shared/` containing all app-agnostic code. This package is consumed by both `packages/excel/` and `packages/powerpoint/` via pnpm workspace protocol.

## What Moves to Shared

### Direct moves (copy as-is, delete from excel)

| Current path | New path | Notes |
|-------------|----------|-------|
| `src/lib/message-utils.ts` | `packages/shared/src/message-utils.ts` | 100% generic |
| `src/lib/truncate.ts` | `packages/shared/src/truncate.ts` | 100% generic |
| `src/lib/sandbox.ts` | `packages/shared/src/sandbox.ts` | 100% generic |
| `src/lib/oauth/index.ts` | `packages/shared/src/oauth/index.ts` | Parameterize storage key |
| `src/lib/skills/index.ts` | `packages/shared/src/skills/index.ts` | 100% generic |
| `src/lib/tools/types.ts` | `packages/shared/src/tools/types.ts` | Remove `DirtyTrackingConfig` (Excel-specific) |
| `src/lib/tools/bash.ts` | `packages/shared/src/tools/bash.ts` | 100% generic |
| `src/lib/tools/read-file.ts` | `packages/shared/src/tools/read-file.ts` | 100% generic |
| `src/taskpane/lockdown.ts` | `packages/shared/src/lockdown.ts` | 100% generic |
| `src/taskpane/index.css` | `packages/shared/src/styles/index.css` | CSS vars + Tailwind base |

### Moves with minor edits

| Current path | New path | Changes needed |
|-------------|----------|----------------|
| `src/lib/provider-config.ts` | `packages/shared/src/provider-config.ts` | Parameterize `STORAGE_KEY` — accept via function arg or make configurable |
| `src/lib/storage/db.ts` | `packages/shared/src/storage/db.ts` | Parameterize DB name + settings key (see below) |
| `src/lib/vfs/index.ts` | `packages/shared/src/vfs/index.ts` | Keep generic; custom commands injected by app |
| `src/taskpane/components/chat/*` | `packages/shared/src/chat/*` | Refactored in [02-refactor-chat-context.md](./02-refactor-chat-context.md) |

### Stays in Excel (NOT shared)

| Path | Why |
|------|-----|
| `src/lib/excel/` | Excel.run API, sheet IDs, tracked context |
| `src/lib/dirty-tracker.ts` | Excel-specific concept |
| `src/lib/tools/` (all except bash, read, types) | Excel-specific tools |
| `src/lib/vfs/custom-commands.ts` | Uses `Excel.run`, `setCellRange`, etc. |

## Storage Parameterization

### db.ts changes

The storage layer needs to accept a config object instead of hardcoding names:

```typescript
// packages/shared/src/storage/db.ts

export interface StorageConfig {
  dbName: string;          // e.g. "OpenExcelDB_v3" or "OpenPptDB_v1"
  dbVersion: number;       // e.g. 30 or 1
  settingsKey: string;     // e.g. "openexcel-workbook-id" or "openppt-document-id"
}

let config: StorageConfig | null = null;

export function initStorage(cfg: StorageConfig) {
  config = cfg;
}
```

- `getOrCreateWorkbookId()` → `getOrCreateDocumentId()` (rename, uses `config.settingsKey`)
- `ChatSession.workbookId` → `ChatSession.documentId` (cosmetic rename)
- `getDb()` uses `config.dbName` and `config.dbVersion`

### provider-config.ts changes

```typescript
// Accept storage key prefix from app
let storageKeyPrefix = "openoffice"; // default

export function initProviderConfig(prefix: string) {
  storageKeyPrefix = prefix;
}
// STORAGE_KEY becomes `${storageKeyPrefix}-provider-config`
```

### oauth/index.ts changes

Same pattern — parameterize `OAUTH_STORAGE_KEY`:

```typescript
let storageKeyPrefix = "openoffice";

export function initOAuth(prefix: string) {
  storageKeyPrefix = prefix;
}
// OAUTH_STORAGE_KEY becomes `${storageKeyPrefix}-oauth-credentials`
```

## VFS Custom Commands

The VFS itself (`index.ts`) is generic. Custom commands are app-specific:

- **Excel** has `csv-to-sheet`, `sheet-to-csv` (call Excel.run)
- **PPT** might have none initially, or could have `pptx-to-text`
- **Generic** commands (`pdf-to-text`, `docx-to-text`, `xlsx-to-csv`, `pdf-to-images`) move to shared

Split `custom-commands.ts`:
- `packages/shared/src/vfs/custom-commands.ts` — `pdfToText`, `pdfToImages`, `docxToText`, `xlsxToCsv`
- `packages/excel/src/lib/vfs/custom-commands.ts` — `csvToSheet`, `sheetToCsv`

The VFS `getCustomCommands()` becomes configurable — app passes its extra commands in.

```typescript
// packages/shared/src/vfs/index.ts
export function getBash(extraCommands?: CustomCommand[]): Bash {
  if (!bash) {
    bash = new Bash({
      fs: getVfs(),
      cwd: "/home/user",
      customCommands: [...getSharedCommands(), ...(extraCommands ?? [])],
    });
  }
  return bash;
}
```

## Package Setup

```json
// packages/shared/package.json
{
  "name": "@open-office/shared",
  "version": "0.0.1",
  "private": true,
  "type": "module",
  "main": "src/index.ts",
  "exports": {
    ".": "./src/index.ts",
    "./chat": "./src/chat/index.ts",
    "./storage": "./src/storage/index.ts",
    "./tools": "./src/tools/index.ts",
    "./vfs": "./src/vfs/index.ts",
    "./oauth": "./src/oauth/index.ts",
    "./skills": "./src/skills/index.ts",
    "./styles": "./src/styles/index.css"
  },
  "dependencies": {
    "@mariozechner/pi-agent-core": "^0.52.6",
    "@mariozechner/pi-ai": "^0.52.6",
    "@sinclair/typebox": "^0.34.48",
    "@streamdown/code": "^1.0.1",
    "idb": "^8.0.3",
    "just-bash": "^2.7.0",
    "lucide-react": "^0.563.0",
    "mammoth": "^1.11.0",
    "pdfjs-dist": "^5.4.624",
    "react": "^18.2.0",
    "react-dom": "^18.2.0",
    "ses": "^1.14.0",
    "streamdown": "^2.1.0",
    "xlsx": "^0.18.5"
  },
  "devDependencies": {
    "@types/react": "^18.2.64",
    "@types/react-dom": "^18.2.21",
    "typescript": "^5.4.2"
  }
}
```

Since shared is source-only (no build step), each app's Vite resolves and bundles it directly via the workspace link. TypeScript path resolution works through the `exports` field.

## Shared Barrel Export

```typescript
// packages/shared/src/index.ts
export * from "./message-utils";
export * from "./truncate";
export * from "./sandbox";
export * from "./lockdown";
export * from "./provider-config";
```

## Checklist

- [ ] Create `packages/shared/` directory structure
- [ ] Create `packages/shared/package.json`
- [ ] Create `packages/shared/tsconfig.json`
- [ ] Move generic files (message-utils, truncate, sandbox, lockdown)
- [ ] Move + parameterize storage (db.ts)
- [ ] Move + parameterize provider-config.ts
- [ ] Move + parameterize oauth/index.ts
- [ ] Move skills/index.ts
- [ ] Move bash + read tools
- [ ] Move tool types (strip DirtyTrackingConfig)
- [ ] Split custom-commands.ts (generic vs Excel-specific)
- [ ] Move CSS (index.css)
- [ ] Create barrel exports
- [ ] Verify all imports resolve
