# VFS + CLI Tools for Agent

## Problem

When the agent needs to move data between uploaded files and the spreadsheet, the naive flow is:

1. LLM calls `read` / `bash cat` to read file content → data enters context window
2. LLM calls `set_cell_range` with all that data → data written out again as tool args

This is extremely token-wasteful. A 1000-row CSV could burn 50k+ tokens round-tripping through the LLM context when the LLM doesn't need to see the data at all.

## Solution

Register custom bash commands via `just-bash`'s `defineCommand` that bridge the VFS and Excel APIs directly. The data flows VFS → Excel (or Excel → VFS) without ever entering the LLM context. The LLM just invokes a short command.

```
User uploads data.csv → VFS
LLM runs: csv-to-sheet data.csv 1 A1    (few tokens)
Data flows: VFS → Excel API              (zero tokens)
```

`just-bash` custom commands receive `(args, ctx)` where `ctx.fs` is the VFS and the execute function is an async closure that can call our Excel APIs directly.

## Architecture

```
┌─────────────────────────────────────────┐
│  LLM Context                            │
│  "csv-to-sheet uploads/data.csv 1 A1"   │  ← few tokens
└──────────────┬──────────────────────────┘
               │ bash tool
               ▼
┌──────────────────────────────────────────┐
│  just-bash                               │
│  ┌─────────────┐    ┌─────────────────┐  │
│  │ Built-in     │    │ Custom commands │  │
│  │ cat,grep,awk │    │ csv-to-sheet   │  │
│  │ sed,jq,sort  │    │ sheet-to-csv   │  │
│  └──────┬──────┘    └───────┬─────────┘  │
│         │                   │            │
│         ▼                   ▼            │
│     ┌───────┐      ┌──────────────┐     │
│     │  VFS  │◄────►│  Excel API   │     │
│     └───────┘      └──────────────┘     │
└──────────────────────────────────────────┘
```

## Commands

### Phase 1: Core data transfer

| Command | Usage | Description |
|---------|-------|-------------|
| `csv-to-sheet` | `csv-to-sheet <file> <sheetId> [startCell]` | Import CSV from VFS into spreadsheet. Parses CSV, calls `setCellRange`. |
| `sheet-to-csv` | `sheet-to-csv <sheetId> <range> [file]` | Export spreadsheet range to CSV file in VFS. Calls `getRangeAsCsv`, writes to VFS. If no file given, writes to stdout (pipeable). |

### Phase 2: More formats & utilities

| Command | Usage | Description |
|---------|-------|-------------|
| `json-to-sheet` | `json-to-sheet <file> <sheetId> [startCell]` | Import JSON array-of-objects. Keys become headers. |
| `sheet-to-json` | `sheet-to-json <sheetId> <range> [file]` | Export range as JSON array-of-objects (first row = keys). |
| `sheet-info` | `sheet-info [sheetId]` | Dump sheet metadata (name, used range, row/col counts) to stdout. Pipeable to `grep`, `jq`. |
| `sheet-ls` | `sheet-ls` | List all sheets with IDs, names, row counts. Like `ls` for the workbook. |

### Phase 3: Advanced

| Command | Usage | Description |
|---------|-------|-------------|
| `sheet-search` | `sheet-search <term> [sheetId]` | Search for text, output matches as TSV. Pipeable. |
| `sheet-formula` | `sheet-formula <sheetId> <range> <formula>` | Apply a formula pattern to a range. |

## Composability Examples

```bash
# Preview before importing
head -5 uploads/sales.csv

# Import CSV
csv-to-sheet uploads/sales.csv 1 A1

# Import only certain columns
cut -d, -f1,3,5 uploads/sales.csv > filtered.csv
csv-to-sheet filtered.csv 1 A1

# Export, transform in bash, re-import to new sheet
sheet-to-csv 1 A1:D100 raw.csv
awk -F, '{print $1,$2*1.1}' OFS=, raw.csv > adjusted.csv
csv-to-sheet adjusted.csv 2 A1

# Pipe export directly to analysis
sheet-to-csv 1 A1:Z1000 | sort -t, -k3 -rn | head -20

# Quick workbook overview
sheet-ls
sheet-info 1 | jq .

# Find and export matching data
sheet-search "error" | cut -f3 > error_cells.txt
```

## Implementation Notes

- Custom commands are registered via `defineCommand` from `just-bash/browser`
- Commands are passed to `new Bash({ customCommands: [...] })` in `src/lib/vfs/index.ts`
- Each command's execute fn is an async closure that calls Excel API functions from `src/lib/excel/api.ts`
- CSV parsing: use `just-bash`'s built-in capabilities or a simple parser (data.csv files are typically well-formed)
- Error handling: return `{ stdout: "", stderr: "error message", exitCode: 1 }` on failure
- System prompt needs updating to document available CLI commands

## File Format Conversion

Binary document formats need external libraries for text extraction. These get wired in as custom bash commands so the agent can extract content without it flowing through the LLM context.

### Libraries

| Format | Library | Browser | Size | Notes |
|--------|---------|---------|------|-------|
| PDF | `pdfjs-dist` | ✅ | ~400KB | Mozilla's PDF.js. Text extraction per page. |
| XLSX/XLS/ODS | `SheetJS` (`xlsx`) | ✅ | ~300KB | Reads all spreadsheet formats. `sheet_to_csv()` for text. |
| DOCX | `mammoth` | ✅ | ~50KB | DOCX → plain text / HTML. Clean output. |

### Already handled by `just-bash` (no extra deps)

- **CSV** — `xan` (headers, select, filter, sort, agg, groupby), `awk`, `cut`, `sort`
- **JSON/YAML/XML/INI/TOML/CSV** — `yq` converts between all of these
- **Plain text** — `cat`, `grep`, `head`, `tail`, etc.

### Commands

| Command | Usage | Description |
|---------|-------|-------------|
| `pdf-to-text` | `pdf-to-text <file> [outfile]` | Extract text from PDF. Stdout if no outfile. |
| `docx-to-text` | `docx-to-text <file> [outfile]` | Extract text from DOCX. |
| `xlsx-to-csv` | `xlsx-to-csv <file> [sheet] [outfile]` | Convert XLSX sheet to CSV in VFS. |
| `xlsx-to-sheet` | `xlsx-to-sheet <file> <sheetId> [startCell]` | Import uploaded XLSX directly into workbook. |

### Coverage

Covers all common "normie" formats:
- **Spreadsheets**: CSV, XLSX, XLS, ODS
- **Documents**: DOCX, PDF
- **Data/config**: JSON, XML, YAML, INI, TOML, CSV
- **Text**: TXT, MD, and anything plain text

Out of scope (rare, not worth bundle size): `.numbers` (Apple), `.pptx`, `.odt`

## Status

- [x] VFS with `just-bash/browser` (InMemoryFs + Bash)
- [x] VFS persistence per chat session (IndexedDB via Dexie)
- [x] File upload UI (paperclip button, file chips)
- [x] `read` tool (returns file content / images to LLM)
- [x] `bash` tool (executes commands in VFS)
- [ ] `csv-to-sheet` custom command
- [ ] `sheet-to-csv` custom command
- [ ] System prompt update for CLI commands
- [ ] Phase 2 commands
- [ ] Phase 3 commands
- [ ] `pdf-to-text` (pdfjs-dist)
- [ ] `docx-to-text` (mammoth)
- [ ] `xlsx-to-csv` / `xlsx-to-sheet` (SheetJS)
