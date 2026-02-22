# open-excel

OpenExcel is an Excel add-in with an AI chat panel. It connects to major LLM providers using your own credentials (BYOK), and can read/write spreadsheets through built-in tools.

https://github.com/user-attachments/assets/50f3ba42-4daa-49d8-b31e-bae9be6e225b

> Reference: Claude for Excel prompt/tool protocol notes are documented at [hewliyang/reversing/claude-for-excel](https://github.com/hewliyang/reversing/tree/main/claude-for-excel).

## Install (End Users)

Download [`manifest.prod.xml`](./manifest.prod.xml), then follow the instructions for your platform:

### Windows
1. **Insert** → **Add-ins** → **My Add-ins**
2. **Upload My Add-in**
3. Select `manifest.prod.xml`
4. Open **Open Excel Chat** from the ribbon

### macOS
1. Copy `manifest.prod.xml` to:
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel
3. **Insert** → **Add-ins** → **My Add-ins**
4. Select **OpenExcel**

### Excel Web
1. Open [excel.office.com](https://excel.office.com)
2. **Insert** → **Add-ins** → **More Add-ins**
3. **Upload My Add-in**
4. Upload `manifest.prod.xml`

## Tools

### Excel tools

| Tool | What it does |
|------|---------------|
| `get_cell_ranges` | Read cell values, formulas, and formats |
| `get_range_as_csv` | Export a range as CSV for analysis |
| `search_data` | Search worksheet data by text |
| `screenshot_range` | Capture a range as an image |
| `get_all_objects` | List tables, charts, pivots, and other objects |
| `set_cell_range` | Write values/formulas/formats to cells |
| `clear_cell_range` | Clear cell contents and/or formatting |
| `copy_to` | Copy ranges with formula translation |
| `modify_sheet_structure` | Insert/delete/hide rows/columns, freeze panes |
| `modify_workbook_structure` | Create/delete/rename/reorder sheets |
| `resize_range` | Resize row heights and column widths |
| `modify_object` | Create/update/delete charts/tables/pivots |
| `eval_officejs` | Run raw Office.js inside Excel.run (sandboxed) |

### File & shell tools

| Tool | What it does |
|------|---------------|
| `read` | Read text files and images from the virtual filesystem |
| `bash` | Run commands in the sandboxed shell |

### Bash custom commands

| Command | What it does |
|---------|---------------|
| `csv-to-sheet` | Import CSV from VFS to a worksheet |
| `sheet-to-csv` | Export worksheet data to CSV |
| `pdf-to-text` | Extract text from PDF files |
| `pdf-to-images` | Render PDF pages to PNG images |
| `docx-to-text` | Extract text from DOCX files |
| `xlsx-to-csv` | Convert uploaded spreadsheet files to CSV |
| `image-to-sheet` | Paint an image into Excel as pixel-art cells |
| `web-search` | Search the web using configured provider |
| `web-fetch` | Fetch web pages/files into VFS |

## Skills

You can install skills from:
- a single `SKILL.md` file, or
- a folder that contains `SKILL.md`.

Manage skills from the Settings tab.

## Providers

- API key (BYOK): OpenAI, Anthropic, Google, Azure, OpenRouter, Groq, xAI, Cerebras, Mistral, etc.
- OAuth: Anthropic (Claude Pro/Max), OpenAI Codex (ChatGPT Plus/Pro)
- Custom endpoint: OpenAI-compatible APIs (Ollama, vLLM, LMStudio, ...)

## Configuration

In **Settings** you can configure:
- Provider, model, and auth method
- CORS proxy
- Thinking level
- Skills
- Web search/fetch providers and API keys

### Web search/fetch credentials

Configure web provider credentials in the Settings UI.

Supported providers:
- DuckDuckGo; search (free, but will rate limit easily)
- Brave; search
- Serper; search
- Exa; search, fetch

More often than not, `basic` fetch is good enough but requires a CORS proxy configured.

## Development

```bash
pnpm install
pnpm start        # start dev server + sideload add-in
pnpm stop         # stop sideloaded add-in
```

## License

MIT
