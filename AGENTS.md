# AGENTS.md

## Project Overview

**OpenExcel** is a Microsoft Office Excel Add-in with an integrated AI chat interface. Users can chat with LLM providers (OpenAI, Anthropic, Google, etc.) directly within Excel using their own API keys (BYOK).

## Tech Stack

- **Framework**: React 18
- **Language**: TypeScript
- **Styling**: Tailwind CSS v4 + CSS variables for theming
- **Icons**: Lucide React (`lucide-react`)
- **Build Tool**: Webpack 5
- **Office Integration**: Office.js API (`@types/office-js`)
- **LLM Integration**: `@mariozechner/pi-ai` (unified LLM API)
- **Dev Server**: webpack-dev-server with HTTPS

## Project Structure

```
open-excel/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── app.tsx              # Root component
│   │   │   └── chat/                # AI Chat UI
│   │   │       ├── index.ts         # Exports
│   │   │       ├── types.ts         # Type definitions
│   │   │       ├── chat-interface.tsx   # Main chat with tabs
│   │   │       ├── chat-context.tsx     # State + pi-ai integration
│   │   │       ├── message-list.tsx     # Message renderer
│   │   │       ├── chat-input.tsx       # Input component
│   │   │       └── settings-panel.tsx   # API key config
│   │   ├── index.tsx                # React entry point
│   │   ├── index.css                # Tailwind + CSS variables
│   │   └── taskpane.html            # HTML template
│   ├── shims/
│   │   └── crypto-shim.js           # Browser crypto polyfill
│   └── commands/
│       └── commands.ts              # Ribbon handlers
├── .plan/                           # Development plans
├── assets/                          # Icons
├── manifest.xml                     # Office Add-in manifest
├── webpack.config.js                # Webpack + browser polyfills
└── package.json
```

## Key Components

### Chat System (`src/taskpane/components/chat/`)

| File                 | Purpose                                                   |
| -------------------- | --------------------------------------------------------- |
| `chat-interface.tsx` | Tab navigation (Chat/Settings), header with clear button  |
| `chat-context.tsx`   | React context, pi-ai streaming, message state, CORS proxy |
| `message-list.tsx`   | Renders user/assistant messages with streaming cursor     |
| `chat-input.tsx`     | Auto-resizing textarea, Enter to send                     |
| `settings-panel.tsx` | Provider/model/API key config, proxy toggle               |

### CSS Variables (Dark Theme)

```css
--chat-font-mono      /* Monospace font stack */
--chat-bg             /* #0a0a0a */
--chat-border         /* #2a2a2a */
--chat-text-primary   /* #e8e8e8 */
--chat-accent         /* #6366f1 (indigo) */
--chat-radius         /* 2px (boxy style) */
```

## LLM Integration

### Supported Providers (via pi-ai)

- OpenAI, Azure OpenAI
- Anthropic (Claude)
- Google (Gemini)
- OpenRouter, Groq, xAI, Cerebras, Mistral

### CORS Proxy

Some providers (Anthropic, Z.ai) require a CORS proxy for browser requests. Users can configure their own proxy URL in settings. The proxy should accept `?url={encodedApiUrl}` format.

### Webpack Browser Polyfills

pi-ai requires Node.js polyfills for browser bundling:

```javascript
// Key webpack.config.js settings
resolve.fallback = { buffer, stream, util, ... }
resolve.alias = { "node:buffer": "buffer/", ... }
DefinePlugin: { "process.versions": "undefined" }  // Prevents Node.js-only imports
```

## Development Commands

```bash
pnpm install             # Install dependencies
pnpm dev-server          # Start dev server (https://localhost:3000)
pnpm start               # Launch Excel with add-in sideloaded
pnpm build               # Production build
pnpm build:dev           # Development build
pnpm lint                # Run Biome linter
pnpm format              # Format code with Biome
pnpm typecheck           # TypeScript type checking
```

## Code Style

- Formatter/linter: Biome
- No JSDoc comments on functions (keep code clean)
- Run `pnpm format` before committing

## Release Workflow

Releases are triggered by pushing a version tag. CI runs quality checks, deploys to Cloudflare Pages, and auto-creates a GitHub release.

```bash
pnpm version patch       # Bump version (patch/minor/major), creates git tag
git push && git push --tags
```

CI workflow (`.github/workflows/release.yml`):
1. Runs typecheck, lint, build
2. Deploys to Cloudflare Pages
3. Creates GitHub release with auto-generated notes

## Configuration Storage

User settings stored in browser localStorage:

| Key                         | Contents                                |
| --------------------------- | --------------------------------------- |
| `openexcel-provider-config` | `{ provider, apiKey, model, useProxy, proxyUrl, thinking }` |

## Excel API Usage

```typescript
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.values = [["value"]];
  await context.sync();
});
```

## Future Development

See `.plan/` directory for roadmap and progress tracking.

### Planned Features

- Excel-specific tools (read/write cells, formulas)
- Markdown rendering in messages
- Code syntax highlighting
- Conversation history/sessions
- Agent with tool execution

## References

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API](https://learn.microsoft.com/en-us/javascript/api/excel)
- [pi-ai GitHub](https://github.com/badlogic/pi-mono)
