# open-excel

An open-source Claude for Excel clone. A Microsoft Office Excel Add-in with an integrated AI chat interface that lets you chat with LLM providers (OpenAI, Anthropic, Google, etc.) directly within Excel using your own API keys (BYOK).

## Installation (End Users)

Download [`manifest.prod.xml`](./manifest.prod.xml) and follow the instructions for your platform:

### Windows

1. Open Excel
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Upload My Add-in**
4. Browse to `manifest.prod.xml` and click OK
5. Click **"Open AI Chat"** in the Home tab ribbon

### macOS

1. Copy `manifest.prod.xml` to:
   ```
   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
   You can do this via Terminal:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   cp manifest.prod.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
2. Quit and reopen Excel
3. Go to **Insert** → **Add-ins** → **My Add-ins**
4. Select "OpenExcel" under **Shared Folder**

### Excel for Web

1. Open a workbook at [excel.office.com](https://excel.office.com)
2. Click **Insert** → **Add-ins** → **More Add-ins**
3. Click **Upload My Add-in**
4. Upload `manifest.prod.xml`

> **Note:** The "Upload My Add-in" option may be disabled by your organization's IT admin.

---

## Development

### Prerequisites

- [Node.js](https://nodejs.org/) (v18 or higher recommended)
- Microsoft Excel (desktop version)
- pnpm (or npm/yarn)

### Setup

```bash
pnpm install
```

### Start Dev Server

This command starts the dev server and sideloads the add-in into Excel:

```bash
pnpm start
```

Excel will launch automatically with the add-in loaded in the taskpane.

### Stop the Add-in

```bash
pnpm stop
```

### Deploy to Production

Builds and deploys to Cloudflare Pages:

```bash
pnpm deploy
```

### Other Commands

| Command | Description |
|---------|-------------|
| `pnpm dev-server` | Start dev server only (https://localhost:3000) |
| `pnpm build` | Production build |
| `pnpm deploy` | Build and deploy to Cloudflare Pages |
| `pnpm lint` | Run linter |
| `pnpm typecheck` | TypeScript type checking |
| `pnpm validate` | Validate the Office manifest |

## Claude for Excel Parity

### Spreadsheet Tools (11)

| Tool | Description |
|------|-------------|
| `get_cell_ranges` | Read cell values, formulas, and formatting |
| `get_range_as_csv` | Pull data as CSV (great for analysis) |
| `search_data` | Find text across the spreadsheet |
| `get_all_objects` | List charts, pivot tables, etc. |
| `set_cell_range` | Write values, formulas, and formatting |
| `clear_cell_range` | Clear cells (content, formatting, or both) |
| `copy_to` | Copy ranges with formula translation |
| `modify_sheet_structure` | Insert/delete/hide/freeze rows/columns |
| `modify_workbook_structure` | Create/delete/rename sheets |
| `resize_range` | Adjust column widths and row heights |
| `modify_object` | Create/update/delete charts and pivot tables |

### Original Tools (1)

| Tool | Description |
|------|-------------|
| `eval_officejs` | Execute arbitrary Office.js code within Excel.run context (escape hatch) |

### Non-Spreadsheet Tools (4)

These are not implemented for obvious reasons. I guess we can do it as BYOK w/ some sandbox & search API providers as well.

| Tool | Description |
|------|-------------|
| `code_execution` | Python with RPC to the sheet (pandas, numpy, etc.) |
| `text_editor_code_execution` | Create/edit files |
| `bash_code_execution` | Run shell commands |
| `web_search` | Search the internet for current info |

## Configuration

On first use, open the Settings tab in the add-in to configure:

1. **Provider** - Select your LLM provider (OpenAI, Anthropic, Google, etc.)
2. **API Key** - Enter your API key for the selected provider
3. **Model** - Choose the model to use

Settings are stored locally in the webview sidecar's localStorage.

## License

MIT
