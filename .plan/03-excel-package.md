# 03 — Excel Package

## Goal

Restructure the existing Excel-specific code into `packages/excel/`, making it a thin consumer of `@open-office/shared`.

## What Lives in packages/excel/

```
packages/excel/
├── package.json
├── tsconfig.json
├── vite.config.ts
├── vitest.config.ts
├── manifest.xml                      # Dev (Host: Workbook, localhost:3000)
├── manifest.prod.xml                 # Prod (openexcel.pages.dev)
├── public/
│   └── assets/                       # Icons (icon-16/32/64/80.png, logo-filled.png)
├── src/
│   ├── lib/
│   │   ├── excel/
│   │   │   ├── api.ts               # Core Excel.run operations (unchanged)
│   │   │   ├── index.ts             # Re-exports
│   │   │   ├── sheet-id-map.ts      # Stable sheet ID tracking (unchanged)
│   │   │   └── tracked-context.ts   # Dirty range tracking proxy (unchanged)
│   │   ├── dirty-tracker.ts         # DirtyRange type + merge/parse helpers (unchanged)
│   │   ├── tools/
│   │   │   ├── index.ts             # EXCEL_TOOLS array (imports from shared for bash/read)
│   │   │   ├── get-cell-ranges.ts   # (unchanged)
│   │   │   ├── get-range-as-csv.ts  # (unchanged)
│   │   │   ├── set-cell-range.ts    # (unchanged)
│   │   │   ├── clear-cell-range.ts  # (unchanged)
│   │   │   ├── copy-to.ts          # (unchanged)
│   │   │   ├── resize-range.ts     # (unchanged)
│   │   │   ├── search-data.ts      # (unchanged)
│   │   │   ├── get-all-objects.ts  # (unchanged)
│   │   │   ├── modify-object.ts    # (unchanged)
│   │   │   ├── modify-sheet-structure.ts  # (unchanged)
│   │   │   ├── modify-workbook-structure.ts # (unchanged)
│   │   │   └── eval-officejs.ts    # (unchanged)
│   │   ├── vfs/
│   │   │   └── custom-commands.ts  # csv-to-sheet, sheet-to-csv (uses Excel.run)
│   │   └── excel-config.ts         # AppConfig for Excel (NEW — see below)
│   ├── taskpane/
│   │   ├── index.tsx               # Entry: renders App
│   │   ├── index.css               # @import "@open-office/shared/styles"
│   │   ├── lockdown.ts             # Re-export from shared? Or keep local
│   │   └── components/
│   │       └── app.tsx             # Creates excelConfig, passes to <ChatInterface config={excelConfig} />
│   ├── commands/
│   │   └── commands.ts             # Ribbon handler (unchanged)
│   ├── shims/
│   │   └── util-types-shim.js      # (unchanged)
│   ├── taskpane.html               # (unchanged)
│   ├── commands.html               # (unchanged)
│   └── global.d.ts                 # (unchanged)
└── tests/
    └── sandbox.test.ts             # (unchanged)
```

## The Excel AppConfig

The main new file — wires everything together:

```typescript
// packages/excel/src/lib/excel-config.ts

import type { AppConfig } from "@open-office/shared/chat";
import { bashTool, readTool } from "@open-office/shared/tools";
import { buildSkillsPromptSection, type SkillMeta } from "@open-office/shared/skills";
import { EXCEL_SPECIFIC_TOOLS } from "./tools";
import { getExcelCustomCommands } from "./vfs/custom-commands";
import { getWorkbookMetadata, navigateTo } from "./excel/api";
import type { DirtyRange } from "./dirty-tracker";

function parseDirtyRanges(result: string): DirtyRange[] | null {
  try {
    const parsed = JSON.parse(result);
    return parsed._dirtyRanges ?? null;
  } catch {
    return null;
  }
}

export const excelConfig: AppConfig = {
  appName: "OpenExcel",

  tools: [readTool, bashTool, ...EXCEL_SPECIFIC_TOOLS],

  buildSystemPrompt: (skills: SkillMeta[]) => {
    return `You are an AI assistant integrated into Microsoft Excel...
    ${buildSkillsPromptSection(skills)}`;
    // (full Excel system prompt here — moved from chat-context.tsx)
  },

  getAppContext: async () => {
    try {
      const metadata = await getWorkbookMetadata();
      return { tag: "wb_context", content: JSON.stringify(metadata, null, 2) };
    } catch {
      return null;
    }
  },

  onToolResult: (toolCallId, resultText) => {
    // Follow mode: navigate to dirty ranges
    const dirtyRanges = parseDirtyRanges(resultText);
    if (dirtyRanges && dirtyRanges.length > 0) {
      const first = dirtyRanges[0];
      if (first.sheetId >= 0) {
        navigateTo(first.sheetId, first.range !== "*" ? first.range : undefined).catch(console.error);
      }
    }
  },

  onCitationClick: (hash) => {
    // Parse #cite:sheetId or #cite:sheetId!A1:B10
    const match = hash.match(/^cite:(\d+)(?:!(.+))?$/);
    if (!match) return false;
    const sheetId = parseInt(match[1], 10);
    const range = match[2];
    navigateTo(sheetId, range).catch(console.error);
    return true;
  },

  extraConfigDefaults: { followMode: true },

  customCommands: getExcelCustomCommands(),

  storage: {
    dbName: "OpenExcelDB_v3",
    dbVersion: 30,
    settingsKey: "openexcel-workbook-id",
    storageKeyPrefix: "openexcel",
  },
};
```

## The Excel App Component

Becomes trivially thin:

```tsx
// packages/excel/src/taskpane/components/app.tsx

import { ChatInterface } from "@open-office/shared/chat";
import { excelConfig } from "../../lib/excel-config";

export default function App() {
  return (
    <div className="h-screen w-full overflow-hidden">
      <ChatInterface config={excelConfig} />
    </div>
  );
}
```

## Excel Tools Index

```typescript
// packages/excel/src/lib/tools/index.ts

// Re-export shared tools
export { bashTool, readTool } from "@open-office/shared/tools";

// Excel-specific tools
export { getCellRangesTool } from "./get-cell-ranges";
export { setCellRangeTool } from "./set-cell-range";
// ... etc

// Combined list (Excel-specific only — bash/read added in excel-config.ts)
export const EXCEL_SPECIFIC_TOOLS = [
  getCellRangesTool,
  getRangeAsCsvTool,
  searchDataTool,
  getAllObjectsTool,
  setCellRangeTool,
  clearCellRangeTool,
  copyToTool,
  modifySheetStructureTool,
  modifyWorkbookStructureTool,
  resizeRangeTool,
  modifyObjectTool,
  evalOfficeJsTool,
];
```

## Excel tool types

The `defineTool` helper in shared won't include `DirtyTrackingConfig`. The Excel package extends it:

```typescript
// packages/excel/src/lib/tools/types.ts

import { defineTool as defineSharedTool } from "@open-office/shared/tools";
import type { DirtyRange } from "../dirty-tracker";

// Re-export shared helpers
export { toolSuccess, toolError, toolText } from "@open-office/shared/tools";

// Extended defineTool with dirty tracking support
export function defineTool<T extends TObject>(config: ToolConfig<T> & {
  dirtyTracking?: { getRanges: (params: Static<T>, result?: unknown) => DirtyRange[] };
}) {
  // Wrap execute to inject _dirtyRanges into results
  // (same logic currently in shared types.ts, just moved here)
}
```

## Custom Commands Split

```typescript
// packages/excel/src/lib/vfs/custom-commands.ts

// Only csv-to-sheet and sheet-to-csv remain here
// (they import from ../excel/api which uses Excel.run)

import type { CustomCommand } from "just-bash/browser";

export function getExcelCustomCommands(): CustomCommand[] {
  return [csvToSheet, sheetToCsv];
}
```

## package.json

```json
{
  "name": "@open-office/excel",
  "version": "0.2.1",
  "private": true,
  "scripts": {
    "build": "vite build",
    "deploy": "pnpm build && pnpm dlx wrangler pages deploy dist --project-name openexcel",
    "dev-server": "vite --mode development",
    "start": "office-addin-debugging start manifest.xml",
    "stop": "office-addin-debugging stop manifest.xml",
    "typecheck": "tsc --noEmit --skipLibCheck",
    "test": "vitest run"
  },
  "dependencies": {
    "@open-office/shared": "workspace:*",
    "@types/office-js": "^1.0.377",
    "@types/office-runtime": "^1.0.35"
  },
  "devDependencies": {
    "@vitejs/plugin-react": "^4.3.4",
    "office-addin-cli": "^2.0.3",
    "office-addin-debugging": "^6.0.3",
    "office-addin-dev-certs": "^2.0.3",
    "office-addin-manifest": "^2.0.3",
    "vite": "^6.3.5",
    "vite-plugin-node-polyfills": "^0.25.0",
    "vite-plugin-static-copy": "^3.2.0",
    "vitest": "^4.0.18"
  }
}
```

Most deps move to shared. Excel keeps only Office.js types and Vite tooling.

## Vite Config

Nearly identical to current, just adjusted paths:

- `root: "src"`
- `publicDir: "../public"`
- `outDir: "../dist"`
- Manifest copy plugin still works
- Alias for `node:util/types` shim stays

## Checklist

- [ ] Create `packages/excel/` directory structure
- [ ] Create `packages/excel/package.json`
- [ ] Create `packages/excel/tsconfig.json` (extends `../../tsconfig.base.json`)
- [ ] Move Excel-specific source files
- [ ] Create `excel-config.ts` with `AppConfig`
- [ ] Update `app.tsx` to use shared `ChatInterface` with config
- [ ] Move Excel system prompt from `chat-context.tsx` to `excel-config.ts`
- [ ] Update tool imports (bash/read from shared, rest local)
- [ ] Split custom commands (keep csv-to-sheet/sheet-to-csv, move rest to shared)
- [ ] Move Vite config, adjust paths
- [ ] Move manifests
- [ ] Move public assets
- [ ] Verify build + dev-server works
- [ ] Verify sideload into Excel works
