# open-excel

An open-source Claude for Excel clone. A Microsoft Office Excel Add-in with an integrated AI chat interface that lets you chat with LLM providers (OpenAI, Anthropic, Google, etc.) directly within Excel using your own API keys (BYOK).

https://github.com/user-attachments/assets/50f3ba42-4daa-49d8-b31e-bae9be6e225b

> **Reference:** The original Claude for Excel system prompt, tools spec, and RPC protocol were reverse-engineered and documented at [hewliyang/reversing/claude-for-excel](https://github.com/hewliyang/reversing/tree/main/claude-for-excel).

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

## Features

### Spreadsheet Tools

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
| `eval_officejs` | Execute arbitrary Office.js code within Excel.run context (escape hatch) |

### File & Shell Tools

| Tool | Description |
|------|-------------|
| `bash` | Execute commands in a sandboxed in-memory shell (pipes, loops, jq, awk, etc.) |
| `read` | Read text files or images from the virtual filesystem |

### CLI Commands

Composable commands available inside the bash shell that bridge uploaded files and Excel:

| Command | Description |
|---------|-------------|
| `csv-to-sheet` | Import a CSV file into a worksheet |
| `sheet-to-csv` | Export a worksheet range to CSV |
| `pdf-to-text` | Extract text from uploaded PDFs |
| `docx-to-text` | Extract text from uploaded Word documents |
| `xlsx-to-csv` | Convert uploaded Excel files to CSV |

### File Uploads

Upload files via the paperclip button or drag-and-drop onto the chat. Files are stored in an in-memory virtual filesystem at `/home/user/uploads/` and persisted per session.

### Skills

Install agent skills to extend the system prompt with specialized instructions. Skills are folders with a `SKILL.md` file containing YAML frontmatter (name + description). Manage skills in the Settings tab.

### Providers & Authentication

Supports all major LLM providers via [pi-ai](https://github.com/badlogic/pi-mono):

- **API Key (BYOK)**: OpenAI, Anthropic, Google, Azure, OpenRouter, Groq, xAI, Cerebras, Mistral
- **OAuth**: Anthropic (Claude Pro/Max), OpenAI Codex (ChatGPT Plus/Pro)
- **Custom Endpoints**: Any OpenAI-compatible API (Ollama, vLLM, LMStudio, etc.)

## Configuration

On first use, open the Settings tab in the add-in to configure:

1. **Provider** — Select a built-in provider or "Custom Endpoint"
2. **Authentication** — Enter an API key, or use OAuth for Anthropic/OpenAI
3. **Model** — Choose from the provider's model list (or enter a model ID for custom endpoints)
4. **CORS Proxy** — Required for Anthropic (OAuth) and Z.ai when running in the browser. You can use the public proxy at `https://proxy.hewliyang.com` — it's [open source](https://github.com/hewliyang/cors-proxy) and doesn't log anything. Otherwise, host your own.
5. **Thinking Level** — Control extended thinking (None/Low/Medium/High)

Settings are stored locally in the browser's localStorage. Session data (messages, uploaded files, skills) is stored in IndexedDB.

---

## Development

### Prerequisites

- [Node.js](https://nodejs.org/) (v18+)
- Microsoft Excel (desktop version)
- pnpm

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

### Other Commands

| Command | Description |
|---------|-------------|
| `pnpm dev-server` | Start dev server only (https://localhost:3000) |
| `pnpm build` | Production build |
| `pnpm deploy` | Build and deploy to Cloudflare Pages |
| `pnpm lint` | Run Biome linter |
| `pnpm format` | Format code with Biome |
| `pnpm typecheck` | TypeScript type checking |
| `pnpm check` | Typecheck + lint |
| `pnpm validate` | Validate the Office manifest |

## License

MIT
