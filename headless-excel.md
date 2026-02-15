# Headless Excel via SpreadJS

Standalone npm package for programmatic Excel workbook control in Node.js — no browser, no Excel, any platform including Linux. Open-excel consumes this for evals, testing, and API server use cases.

## How it works

[SpreadJS](https://developer.mescius.com/spreadjs) is a commercial JavaScript spreadsheet engine (formulas, charts, tables, pivot tables, formatting, xlsx import/export). It's designed for browsers but runs in Node.js with DOM shims.

**Dependencies:**
```bash
# System (Linux)
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev

# NPM
npm install @mescius/spread-sheets @mescius/spread-sheets-io @mescius/spread-sheets-charts @mescius/spread-sheets-pivot-addon happy-dom canvas
```

**Setup (~30 lines of shims):**
```js
const { Window } = require('happy-dom');
const canvas = require('canvas');
const win = new Window({ url: 'http://localhost' });

global.self = global;
global.window = win;
global.document = win.document;
global.navigator = win.navigator;
global.HTMLCollection = win.HTMLCollection;
global.getComputedStyle = win.getComputedStyle.bind(win);
global.customElements = win.customElements;
global.HTMLElement = win.HTMLElement;
global.HTMLDivElement = win.HTMLDivElement;
global.HTMLCanvasElement = win.HTMLCanvasElement;
global.HTMLImageElement = win.HTMLImageElement;
global.Image = win.Image;
global.Event = win.Event;
global.MouseEvent = win.MouseEvent;
global.KeyboardEvent = win.KeyboardEvent;
global.PointerEvent = win.PointerEvent || win.MouseEvent;
global.TouchEvent = win.TouchEvent || class TouchEvent {};
global.WheelEvent = win.WheelEvent || win.Event;
global.MutationObserver = win.MutationObserver;
global.ResizeObserver = class ResizeObserver { observe(){} disconnect(){} unobserve(){} };
global.requestAnimationFrame = (cb) => setTimeout(cb, 0);
global.cancelAnimationFrame = (id) => clearTimeout(id);
// Custom FileReader — happy-dom's breaks SpreadJS xlsx import
global.FileReader = class NodeFileReader {
  constructor() { this.result = null; this.error = null; this.onload = null; this.onerror = null; this.onloadend = null; this.readyState = 0; }
  _done(r) { this.readyState = 2; this.result = r; this.onload?.({ target: this }); this.onloadend?.({ target: this }); }
  _fail(e) { this.error = e; this.onerror?.({ target: this }); this.onloadend?.({ target: this }); }
  readAsArrayBuffer(input) { this.readyState = 1; if (input?.arrayBuffer) input.arrayBuffer().then(ab => this._done(ab)).catch(e => this._fail(e)); else if (input instanceof ArrayBuffer) process.nextTick(() => this._done(input)); }
  readAsBinaryString(input) { this.readyState = 1; const f = (ab) => { const a = new Uint8Array(ab); let s=''; for(let i=0;i<a.length;i++) s+=String.fromCharCode(a[i]); return s; }; if (input?.arrayBuffer) input.arrayBuffer().then(ab => this._done(f(ab))).catch(e => this._fail(e)); }
  readAsDataURL(input) { this.readyState = 1; if (input?.arrayBuffer) input.arrayBuffer().then(ab => this._done('data:'+(input.type||'application/octet-stream')+';base64,'+Buffer.from(ab).toString('base64'))).catch(e => this._fail(e)); }
  readAsText(input, enc) { this.readyState = 1; if (input?.arrayBuffer) input.arrayBuffer().then(ab => this._done(new TextDecoder(enc||'utf-8').decode(ab))).catch(e => this._fail(e)); }
  abort() {} addEventListener(e, fn) { this['on'+e] = fn; } removeEventListener() {}
};
global.DOMParser = win.DOMParser;
global.XMLSerializer = win.XMLSerializer;
global.canvas = canvas;
global.devicePixelRatio = 1;
global.location = win.location;
global.innerWidth = 800;
global.innerHeight = 600;
global.addEventListener = () => {};
global.removeEventListener = () => {};
global.getSelection = () => ({ removeAllRanges: () => {}, addRange: () => {} });

var GC = require('@mescius/spread-sheets');
require('@mescius/spread-sheets-io');        // xlsx I/O
require('@mescius/spread-sheets-charts');     // charts
require('@mescius/spread-sheets-pivot-addon'); // pivot tables
var spread = new GC.Spread.Sheets.Workbook();
```

**XLSX I/O:**
```js
// Save
spread.save((blob) => {
  blob.arrayBuffer().then(buf => fs.writeFileSync('out.xlsx', Buffer.from(buf)));
}, (err) => console.error(err), { fileType: GC.Spread.Sheets.FileType.excel });

// Load
const buf = fs.readFileSync('input.xlsx');
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
spread.import(ab, () => {
  // Trial adds an "Evaluation Version" watermark sheet at index 0 — strip it
  for (let i = spread.getSheetCount() - 1; i >= 0; i--) {
    if (spread.getSheet(i).name() === 'Evaluation Version') spread.removeSheet(i);
  }
  // Ready to use
}, (err) => console.error(err), { fileType: GC.Spread.Sheets.FileType.excel });
```

## Verified working

- Values: `sheet.setValue(row, col, val)` / `sheet.getValue(row, col)`, `setArray` / bulk data
- Formulas: SUM, AVERAGE, MAX, MIN, COUNT, UPPER, LEN, IF, CONCATENATE, VLOOKUP (500+ functions, auto-calc)
- Sheet operations: add (must pass `Worksheet` instance), rename, multi-sheet, cross-sheet formula refs
- Cell styling: backColor, foreColor, font
- Row/column sizing: setColumnWidth, setRowHeight
- Range API: getRange with row/col/count
- Number formatting: `setFormatter` / `getText`
- Cell merging: `addSpan` / `getSpan`
- JSON serialization: `spread.toJSON()` / `spread.fromJSON()` roundtrip (preserves formulas, styles, charts, pivots)
- **Pivot tables** (`@mescius/spread-sheets-pivot-addon`): create, add row/column/value/filter fields, field manipulation, remove fields, JSON roundtrip
- **Charts** (`@mescius/spread-sheets-charts`): column, line, pie, bar, area chart types; title, legend, series, axes config; add/remove/modify charts; change chart type; update data range; JSON roundtrip
- **Tables**: create with themes, row/column count, totals row (`showFooter`)
- **XLSX I/O** (`@mescius/spread-sheets-io`): save to xlsx, load from xlsx, full roundtrip (save → load → modify → save → load). Values, formulas, styles, tables, charts, pivots all survive roundtrip



## Integration with open-excel

All Excel tools funnel through `src/lib/excel/api.ts`. The tools layer (`src/lib/tools/*.ts`) never touches `Excel.run` directly (except `eval-officejs.ts`).

To run headless:
1. Create `src/lib/excel/api-spreadjs.ts` implementing the same exports as `api.ts`
2. Toggle via env: `EXCEL_BACKEND=spreadjs|officejs`
3. SpreadJS API is synchronous — no `context.load()` / `context.sync()` dance

Key API mapping:
| Office.js | SpreadJS |
|---|---|
| `Excel.run(async (ctx) => ...)` | Direct calls (synchronous) |
| `range.load('values'); await ctx.sync()` | `sheet.getValue(row, col)` |
| `range.values = [[...]]` | `sheet.setValue(row, col, val)` |
| `worksheet.getRange("A1:C10")` | `sheet.getRange(0, 0, 10, 3)` |
| `workbook.worksheets.add("name")` | `spread.addSheet(index, new Worksheet(name))` |

## SpreadJS quirks

- `spread.addSheet(index)` returns `undefined` — must pass a `new GC.Spread.Sheets.Worksheet(name)` as second arg
- Table names can't look like cell references (e.g., `T1` fails, `Table1` works) — same rule as Excel itself
- `chart.dataRange()` returns fully-qualified range (`Sheet1!$A$1:$B$7`) even if set without sheet prefix

## Limitations

- `eval_officejs` tool runs raw Office.js — would need a SpreadJS equivalent or be disabled
- SpreadJS is commercial (~$2k license), installed as a **peer dependency** (BYOL). Trial works without a key — on import, an "Evaluation Version" watermark sheet is prepended (stripped automatically by the package). Users can pass a license key via `init({ licenseKey })` to avoid the watermark entirely
- `canvas` npm package requires native libs (Cairo/Pango) — standard on most Linux CI images. **Canvas is required** — even pure data operations (tables, etc.) trigger internal rendering code that crashes without it

## Package design

Minimal wrapper — the package's job is shim setup + I/O helpers, not re-wrapping SpreadJS's API.

**Public API:**
```js
import { init } from 'headless-excel';

// With license (no eval watermark on import)
const { GC, Workbook } = init({ licenseKey: 'xxx' });

// Without license (eval sheet auto-stripped on import)
const { GC, Workbook } = init();

// Create from scratch
const wb = new Workbook();
const sheet = wb.getActiveSheet();
sheet.setValue(0, 0, 'Hello');
await wb.save('output.xlsx');

// Load existing
const wb2 = await Workbook.open('input.xlsx');

// Access raw SpreadJS API
const spread = wb2.spread; // GC.Spread.Sheets.Workbook instance
```

**What the package does internally:**
1. Sets up happy-dom + canvas globals (shim layer)
2. Provides custom `FileReader` (happy-dom's breaks SpreadJS xlsx import)
3. Sets `GC.Spread.Sheets.LicenseKey` if provided
4. Strips "Evaluation Version" watermark sheet on import (trial mode)
5. Exposes `GC.Spread.Sheets` namespace for direct SpreadJS access
6. Provides `dispose()` for cleanup (`window.close()` to prevent happy-dom leaks)

**What consumers do:**
- Use SpreadJS API directly for all spreadsheet operations
- Build higher-level abstractions (like open-excel's `lib/tools/*`) on top

**Peer dependencies** (user installs separately):
- `@mescius/spread-sheets`
- `@mescius/spread-sheets-io`
- `@mescius/spread-sheets-charts` (optional)
- `@mescius/spread-sheets-pivot-addon` (optional)

**Dependencies** (bundled):
- `happy-dom`
- `canvas`

**Considerations:**
- happy-dom `Window` leaks if not closed — `dispose()` must call `window.close()`
- Concurrency: SpreadJS uses globals (`window`, `document`) — multiple workbooks in one process share shims (probably fine, needs testing; if not, worker threads per workbook)
- `canvas` requires Cairo/Pango system libs on Linux — document in README, consider providing Docker snippet

## Use cases

- **Generic**: any Node.js app that needs to create/read/modify xlsx files with full Excel fidelity (formulas, charts, pivots, tables)
- **open-excel evals** (e.g. ib-bench): load xlsx → run agent → score output, all on Linux CI
- **open-excel testing**: run tool integration tests without Excel
- **API server**: expose the agent as an HTTP API backed by SpreadJS
