# Monorepo Migration: open-excel → open-office

## Goal

Restructure the `open-excel` repo into a **pnpm workspace monorepo** that houses both `open-excel` and `open-ppt` (and future Office add-ins), sharing ~70% of the codebase via a `shared` package.

## Repo Naming

- GitHub repo: rename `hewliyang/open-excel` → `hewliyang/open-office`
- Git remote updates after rename
- Each app keeps its own identity (`OpenExcel`, `OpenPPT`)

## Target Structure

```
open-office/
├── pnpm-workspace.yaml
├── package.json                     # Root: scripts, shared devDeps
├── biome.json                       # Shared linter config
├── tsconfig.base.json               # Shared TS config
├── .github/workflows/
│   ├── ci.yml                       # Runs checks on all packages
│   └── release-excel.yml            # Tag-triggered deploy for Excel
│   └── release-ppt.yml              # Tag-triggered deploy for PPT
├── packages/
│   ├── shared/                      # @open-office/shared
│   │   ├── package.json
│   │   ├── tsconfig.json
│   │   └── src/
│   │       ├── chat/                # Chat UI components (React)
│   │       ├── storage/             # IndexedDB (sessions, VFS, skills)
│   │       ├── vfs/                 # Virtual filesystem + bash (generic parts)
│   │       ├── sandbox/             # SES lockdown + compartment
│   │       ├── skills/              # Skill install/management
│   │       ├── tools/               # bash, read, defineTool helpers
│   │       ├── oauth/               # OAuth PKCE flow
│   │       ├── provider-config/     # Provider selection, model config
│   │       ├── message-utils.ts     # AgentMessage → ChatMessage, stats
│   │       └── truncate.ts          # Output truncation
│   │
│   ├── excel/                       # @open-office/excel (current open-excel)
│   │   ├── package.json
│   │   ├── tsconfig.json
│   │   ├── vite.config.ts
│   │   ├── manifest.xml             # Dev manifest (Host: Workbook)
│   │   ├── manifest.prod.xml        # Prod manifest
│   │   └── src/
│   │       ├── lib/
│   │       │   ├── excel/           # Excel.run API, sheet IDs, tracked context
│   │       │   ├── tools/           # 14 Excel-specific tools + EXCEL_TOOLS registry
│   │       │   ├── dirty-tracker.ts # Excel dirty range tracking
│   │       │   └── vfs/             # custom-commands.ts (csv-to-sheet, sheet-to-csv)
│   │       ├── taskpane/
│   │       │   ├── index.tsx        # Entry: title = "OpenExcel"
│   │       │   ├── index.css        # Tailwind + CSS vars
│   │       │   ├── lockdown.ts      # SES lockdown
│   │       │   └── components/
│   │       │       └── app.tsx      # Thin wrapper: passes Excel config to shared ChatInterface
│   │       ├── commands/            # Ribbon command handlers
│   │       ├── taskpane.html
│   │       ├── commands.html
│   │       └── global.d.ts
│   │
│   └── powerpoint/                  # @open-office/powerpoint (new)
│       ├── package.json
│       ├── tsconfig.json
│       ├── vite.config.ts
│       ├── manifest.xml             # Dev manifest (Host: Presentation)
│       ├── manifest.prod.xml        # Prod manifest
│       └── src/
│           ├── lib/
│           │   ├── ppt/             # withSlideZip engine, slide master cleanup, XML utils
│           │   ├── tools/           # 9 PPT tools + PPT_TOOLS registry
│           │   └── vfs/             # custom-commands.ts (PPT-specific, if any)
│           ├── taskpane/
│           │   ├── index.tsx        # Entry: title = "OpenPPT"
│           │   ├── index.css
│           │   ├── lockdown.ts
│           │   └── components/
│           │       └── app.tsx      # Thin wrapper: passes PPT config to shared ChatInterface
│           ├── commands/
│           ├── taskpane.html
│           ├── commands.html
│           └── global.d.ts
```

## Plan Documents

| Plan | What |
|------|------|
| [01-shared-package.md](./01-shared-package.md) | Extract shared code into `packages/shared/` |
| [02-refactor-chat-context.md](./02-refactor-chat-context.md) | Make ChatProvider app-agnostic via config |
| [03-excel-package.md](./03-excel-package.md) | Restructure existing Excel code into `packages/excel/` |
| [04-powerpoint-package.md](./04-powerpoint-package.md) | Build the PowerPoint add-in from reversed specs |
| [05-infra.md](./05-infra.md) | Root config, CI/CD, repo rename, Cloudflare |

## Execution Order

1. **Phase 1 — Scaffolding** (01, 05): Set up monorepo root, pnpm workspace, shared package skeleton
2. **Phase 2 — Extract shared** (01, 02): Move generic code to shared, refactor ChatProvider
3. **Phase 3 — Excel package** (03): Wire up Excel as a thin consumer of shared
4. **Phase 4 — Verify Excel** : Confirm open-excel still builds, dev-server works, sideloads correctly
5. **Phase 5 — PowerPoint package** (04): Implement PPT tools, system prompt, manifest
6. **Phase 6 — Infra** (05): CI/CD, repo rename, Cloudflare Pages projects

## Complexity Estimate

- Phase 1-3: ~1 day (mostly moving files, small refactors)
- Phase 4: ~0.5 day (verify nothing broke)
- Phase 5: ~1.5 days (PPT tools are already reversed, main work is wiring up)
- Phase 6: ~0.5 day (CI/CD, rename)
- **Total: ~3.5 days**

## Key Decisions

1. **No npm publishing** for shared — use workspace protocol (`workspace:*`)
2. **Shared is source-only** — each app's Vite bundles shared directly (no separate build step)
3. **Separate Cloudflare Pages projects** — `openexcel` (exists) + `openppt` (new)
4. **Separate localStorage keys** — `openexcel-*` stays for Excel, `openppt-*` for PPT
5. **Shared IDB or separate?** — Separate DBs (`OpenExcelDB_v3`, `OpenPptDB_v1`). Apps are different browser origins on Cloudflare anyway, so this is automatic.
6. **`followMode`** stays Excel-specific (navigates to dirty cells). PPT won't have this initially.
