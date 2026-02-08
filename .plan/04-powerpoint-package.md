# 04 — PowerPoint Package

## Goal

Build `packages/powerpoint/` — the new `open-ppt` add-in, using reversed specifications from `reversing/claude-for-powerpoint/`.

## Source Material

From `reversing/claude-for-powerpoint/`:
- `system_prompt.md` — complete system prompt (~800 lines)
- `tools.json` — 10 tool definitions (JSON schema)
- `reversed-tools.ts` — full tool implementations (~900 lines)
- `convo.json` — example conversation for reference

## Package Structure

```
packages/powerpoint/
├── package.json
├── tsconfig.json
├── vite.config.ts
├── manifest.xml                      # Dev (Host: Presentation, localhost:3001)
├── manifest.prod.xml                 # Prod (openppt.pages.dev)
├── public/
│   └── assets/                       # Icons
├── src/
│   ├── lib/
│   │   ├── ppt/
│   │   │   ├── slide-zip.ts         # withSlideZip engine (core export/import)
│   │   │   ├── slide-master.ts      # cleanupSlideMasters, forceRemoveMaster
│   │   │   ├── xml-utils.ts         # escapeXml, sanitizeXmlAmpersands, findShapeByName
│   │   │   ├── security.ts          # extractExternalReferences, blocked methods
│   │   │   ├── timeout.ts           # createTimeoutGuard, TimeoutError
│   │   │   ├── office-run.ts        # safeOfficeRun (OfficeOnline serialization)
│   │   │   └── index.ts             # Re-exports
│   │   ├── tools/
│   │   │   ├── index.ts             # PPT_TOOLS array
│   │   │   ├── execute-office-js.ts # execute_office_js (SES sandbox + context proxy)
│   │   │   ├── edit-slide-xml.ts    # edit_slide_xml / edit_slide_chart (same impl)
│   │   │   ├── edit-slide-master.ts # edit_slide_master
│   │   │   ├── edit-slide-text.ts   # edit_slide_text
│   │   │   ├── read-slide-text.ts   # read_slide_text
│   │   │   ├── screenshot-slide.ts  # screenshot_slide
│   │   │   ├── duplicate-slide.ts   # duplicate_slide
│   │   │   ├── verify-slides.ts     # verify_slides
│   │   │   └── sandbox.ts           # compileSlideCode, compileOfficeJsCode, context proxy
│   │   └── ppt-config.ts            # AppConfig for PowerPoint
│   ├── taskpane/
│   │   ├── index.tsx                # Entry: title = "OpenPPT"
│   │   ├── index.css                # @import "@open-office/shared/styles"
│   │   └── components/
│   │       └── app.tsx              # <ChatInterface config={pptConfig} />
│   ├── commands/
│   │   └── commands.ts
│   ├── taskpane.html
│   ├── commands.html
│   └── global.d.ts
```

## Tools to Implement

From `tools.json` and `reversed-tools.ts`:

| Tool | Description | Complexity |
|------|-------------|-----------|
| `execute_office_js` | Arbitrary Office.js in PPT context | High — needs SES sandbox + permission proxy |
| `edit_slide_xml` | Raw OOXML manipulation via JSZip | Medium — uses withSlideZip engine |
| `edit_slide_chart` | Same impl as edit_slide_xml | Low — alias |
| `edit_slide_master` | Edit slide master OOXML + cleanup | Medium — withSlideZip + cleanupSlideMasters |
| `edit_slide_text` | Replace paragraph XML in a shape | Medium — XML parsing/replacement |
| `read_slide_text` | Read raw `<a:p>` XML from shape | Medium — withSlideZip read-only |
| `screenshot_slide` | Get slide as base64 PNG | Low — single API call |
| `duplicate_slide` | Export + re-import slide | Low |
| `verify_slides` | Check overlaps + overflows | Medium — shape geometry math |

## Implementation Plan

### 1. Core Engine (`src/lib/ppt/`)

Port from `reversed-tools.ts`, split into focused modules:

**`slide-zip.ts`** — the `withSlideZip<T>()` function:
- Export slide as base64 → JSZip
- Run callback with `{ zip, markDirty }`
- If dirty: sanitize XML, check external refs, re-import via `insertSlidesFromBase64`
- Restore slide selection
- This is the foundation for 5 of the 9 tools

**`xml-utils.ts`**:
- `escapeXml()` — escape &, <, >, ", '
- `sanitizeXmlAmpersands()` — fix bare `&` in XML
- `findShapeByName()` — locate shape by name + occurrence in OOXML DOM

**`security.ts`**:
- `extractExternalReferences()` — scan .rels files for external targets
- `BLOCKED_CONTEXT_METHODS` — insertSlidesFromBase64, etc.
- `BLOCKED_OFFICE_PATHS` — context.ui.openBrowserWindow, etc.

**`timeout.ts`**:
- `createTimeoutGuard()` — wall-clock timeout with pause/resume for permission prompts

**`office-run.ts`**:
- `safeOfficeRun()` — serialize Office.run calls on OfficeOnline + add timeout

**`slide-master.ts`**:
- `cleanupSlideMasters()` — reassign layouts, remove orphaned masters

### 2. Sandbox (`src/lib/tools/sandbox.ts`)

The PPT sandbox is more sophisticated than Excel's:
- `compileSlideCode()` — for OOXML tools (edit_slide_xml, etc.). Globals: DOMParser, XMLSerializer, escapeXml, Math, Date
- `compileOfficeJsCode()` — for execute_office_js. Adds proxied Office/PowerPoint globals + `pptx.withSlideZip` helper
- `createContextProxy()` — deep proxy that tracks mutations + prompts user on sync()
- `createBlockingProxy()` — blocks nested paths on Office globals

Key difference from Excel: the PPT sandbox exposes `pptx.withSlideZip` inside `execute_office_js` so the LLM can do OOXML manipulation from within Office.js code.

### 3. Tools (`src/lib/tools/`)

Each tool uses `defineTool` from shared (no dirty tracking needed for PPT):

```typescript
// Example: screenshot_slide
export const screenshotSlideTool = defineTool({
  name: "screenshot_slide",
  label: "Screenshot Slide",
  description: "Take a screenshot of a slide...",
  parameters: Type.Object({
    slide_index: Type.Number({ description: "0-based slide index" }),
    explanation: Type.Optional(Type.String({ maxLength: 50 })),
  }),
  execute: async (_id, params) => {
    const imageData = await safeOfficeRun(
      PowerPoint.run.bind(PowerPoint),
      async (context) => {
        const result = context.presentation.slides
          .getItemAt(params.slide_index)
          .getImageAsBase64({ width: 960 });
        await context.sync();
        return result.value;
      }
    );
    return {
      content: [
        { type: "text", text: `Screenshot of slide ${params.slide_index + 1}` },
        { type: "image", data: imageData, mimeType: "image/png" },
      ],
      details: undefined,
    };
  },
});
```

### 4. App Config (`src/lib/ppt-config.ts`)

```typescript
export const pptConfig: AppConfig = {
  appName: "OpenPPT",

  tools: [readTool, bashTool, ...PPT_TOOLS],

  buildSystemPrompt: (skills) => {
    // The reversed system prompt from system_prompt.md
    // + initial_state placeholder (filled by getAppContext)
    return `${SYSTEM_PROMPT}\n${buildSkillsPromptSection(skills)}`;
  },

  getAppContext: async () => {
    // Get presentation state: slide count, masters, dimensions, theme info
    const state = await getPresentationState();
    return { tag: "initial_state", content: JSON.stringify(state, null, 2) };
  },

  // No follow mode for PPT (no equivalent of navigateTo)
  onToolResult: undefined,
  onCitationClick: undefined,

  customCommands: [],   // No PPT-specific bash commands initially

  storage: {
    dbName: "OpenPptDB_v1",
    dbVersion: 1,
    settingsKey: "openppt-document-id",
    storageKeyPrefix: "openppt",
  },
};
```

### 5. Initial State

The system prompt references `<initial_state>` — presentation metadata injected per-message:

```typescript
async function getPresentationState(): Promise<object> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const pageSetup = context.presentation.pageSetup;
    const masters = context.presentation.slideMasters;

    slides.load("items");
    pageSetup.load(["slideWidth", "slideHeight"]);
    masters.load("items");
    await context.sync();

    // Load layouts for each master
    for (const master of masters.items) {
      master.layouts.load("items/name,items/id");
    }
    await context.sync();

    // Detect if default theme
    // Check if content exists on any slide
    // Build masters array with layout names/IDs

    return {
      slideCount: slides.items.length,
      slideWidth: pageSetup.slideWidth,
      slideHeight: pageSetup.slideHeight,
      isDefaultTheme: /* detect */,
      hasContent: /* detect */,
      masters: masters.items.map(m => ({
        layouts: m.layouts.items.map(l => ({ name: l.name, id: l.id }))
      })),
    };
  });
}
```

## Manifest

```xml
<!-- Key differences from Excel manifest -->
<Hosts>
  <Host Name="Presentation"/>  <!-- instead of "Workbook" -->
</Hosts>
<!-- Different add-in ID (new UUID) -->
<!-- DisplayName: "OpenPPT" / "OpenPPT (Dev)" -->
<!-- URLs: openppt.pages.dev / localhost:3001 -->
```

Use port 3001 for PPT dev server to avoid conflict with Excel's 3000.

## Dependencies

```json
{
  "dependencies": {
    "@open-office/shared": "workspace:*",
    "@types/office-js": "^1.0.377",
    "jszip": "^3.10.1"
  }
}
```

PPT needs `jszip` for PPTX manipulation (Excel didn't need it). Shared already has most deps.

## Checklist

- [ ] Create `packages/powerpoint/` directory structure
- [ ] Create package.json, tsconfig.json, vite.config.ts
- [ ] Create manifests (dev + prod)
- [ ] Implement core engine (slide-zip.ts, xml-utils.ts, security.ts, timeout.ts, office-run.ts)
- [ ] Implement sandbox (compileSlideCode, compileOfficeJsCode, context proxy)
- [ ] Implement all 9 tools
- [ ] Port system prompt from reversed `system_prompt.md`
- [ ] Implement `getPresentationState()` for initial_state
- [ ] Create `ppt-config.ts` (AppConfig)
- [ ] Create app.tsx, index.tsx, index.css
- [ ] Create commands.ts, commands.html, taskpane.html
- [ ] Create public/assets icons
- [ ] Verify build
- [ ] Test sideload into PowerPoint
