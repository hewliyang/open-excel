# open-excel

An open-source Claude for Excel clone. A Microsoft Office Excel Add-in with an integrated AI chat interface that lets you chat with LLM providers (OpenAI, Anthropic, Google, etc.) directly within Excel using your own API keys (BYOK).

## Prerequisites

- [Node.js](https://nodejs.org/) (v18 or higher recommended)
- Microsoft Excel (desktop version)
- npm or yarn

## Installation

```bash
npm install
```

## Development

### Start the Add-in

This command starts the dev server and sideloads the add-in into Excel:

```bash
npm run start
```

Excel will launch automatically with the add-in loaded in the taskpane.

### Stop the Add-in

To stop debugging and unload the add-in:

```bash
npm run stop
```

### Other Useful Commands

| Command | Description |
|---------|-------------|
| `npm run dev-server` | Start the dev server only (https://localhost:3000) |
| `npm run build` | Production build |
| `npm run build:dev` | Development build |
| `npm run watch` | Watch mode for development |
| `npm run lint` | Run linter |
| `npm run lint:fix` | Fix linting issues |
| `npm run typecheck` | TypeScript type checking |
| `npm run validate` | Validate the Office manifest |

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
