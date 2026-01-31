/* global Excel */

export interface CellData {
  value: string | number | boolean | null;
  formula?: string;
}

export interface CellStyle {
  sz?: number;
  color?: string;
  family?: string;
  fgColor?: string;
  bold?: boolean;
  italic?: boolean;
}

export interface WorksheetInfo {
  name: string;
  sheetId: number;
  dimension: string;
  cells: Record<string, string | number | boolean | null>;
  formulas?: Record<string, string>;
  styles?: Record<string, CellStyle>;
  borders?: Record<string, unknown>;
}

export interface GetCellRangesResult {
  success: boolean;
  hasMore: boolean;
  worksheet: WorksheetInfo;
}

function columnIndexToLetter(index: number): string {
  let letter = "";
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function cellAddress(rowIndex: number, colIndex: number): string {
  return `${columnIndexToLetter(colIndex)}${rowIndex + 1}`;
}

function parseRangeAddress(address: string): { startCol: number; startRow: number } {
  const clean = address.split("!").pop()?.split(":")[0] || "A1";
  const match = clean.match(/([A-Z]+)(\d+)/);
  if (!match) return { startCol: 0, startRow: 0 };
  const col = match[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number.parseInt(match[2], 10) - 1;
  return { startCol: col, startRow: row };
}

function excelColorToHex(color: Excel.RangeFont | Excel.RangeFill): string | undefined {
  const c = color as { color?: string };
  if (!c.color || c.color === "null") return undefined;
  if (c.color.startsWith("#")) return c.color.toUpperCase();
  return c.color;
}

export async function getWorksheetById(
  context: Excel.RequestContext,
  sheetId: number
): Promise<Excel.Worksheet | null> {
  const sheets = context.workbook.worksheets;
  sheets.load("items");
  await context.sync();

  for (const sheet of sheets.items) {
    sheet.load("id,name");
  }
  await context.sync();

  for (const sheet of sheets.items) {
    const numericId = Number.parseInt(sheet.id.replace(/\D/g, ""), 10);
    if (numericId === sheetId || sheet.id === String(sheetId)) {
      return sheet;
    }
  }

  if (sheetId >= 1 && sheetId <= sheets.items.length) {
    return sheets.items[sheetId - 1];
  }

  return null;
}

export async function getCellRanges(
  sheetId: number,
  ranges: string[],
  options: { includeStyles?: boolean; cellLimit?: number } = {}
): Promise<GetCellRangesResult> {
  const { includeStyles = true, cellLimit = 2000 } = options;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) {
      throw new Error(`Worksheet with ID ${sheetId} not found`);
    }

    sheet.load("name,id");
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();

    const dimension = usedRange.isNullObject ? "A1" : usedRange.address.split("!")[1] || "A1";

    const cells: Record<string, string | number | boolean | null> = {};
    const formulas: Record<string, string> = {};
    const styles: Record<string, CellStyle> = {};
    let totalCells = 0;
    let hasMore = false;

    for (const rangeAddr of ranges) {
      if (totalCells >= cellLimit) {
        hasMore = true;
        break;
      }

      const range = sheet.getRange(rangeAddr);
      range.load("values,formulas,address,rowCount,columnCount");

      if (includeStyles) {
        range.load("format/font,format/fill");
        range.format.font.load("name,size,color,bold,italic");
        range.format.fill.load("color");
      }

      await context.sync();

      const startAddress = range.address.split("!")[1]?.split(":")[0] || "A1";
      const startMatch = startAddress.match(/([A-Z]+)(\d+)/);
      const startCol = startMatch ? startMatch[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1 : 0;
      const startRow = startMatch ? Number.parseInt(startMatch[2], 10) - 1 : 0;

      for (let r = 0; r < range.rowCount && totalCells < cellLimit; r++) {
        for (let c = 0; c < range.columnCount && totalCells < cellLimit; c++) {
          const addr = cellAddress(startRow + r, startCol + c);
          const value = range.values[r][c];
          const formula = range.formulas[r][c];

          if (value !== null && value !== "" && value !== undefined) {
            cells[addr] = value as string | number | boolean;
            totalCells++;
          }

          if (typeof formula === "string" && formula.startsWith("=")) {
            formulas[addr] = formula;
          }
        }
      }

      if (includeStyles && range.format) {
        const rangeStyle: CellStyle = {};
        const font = range.format.font;
        const fill = range.format.fill;

        if (font.size) rangeStyle.sz = font.size;
        if (font.name) rangeStyle.family = font.name;
        if (font.bold) rangeStyle.bold = font.bold;
        if (font.italic) rangeStyle.italic = font.italic;

        const fontColor = excelColorToHex(font as Excel.RangeFont);
        if (fontColor) rangeStyle.color = fontColor;

        const fillColor = excelColorToHex(fill as Excel.RangeFill);
        if (fillColor) rangeStyle.fgColor = fillColor;

        if (Object.keys(rangeStyle).length > 0) {
          styles[rangeAddr] = rangeStyle;
        }
      }
    }

    const sheetNumericId = Number.parseInt(sheet.id.replace(/\D/g, ""), 10) || sheetId;

    return {
      success: true,
      hasMore,
      worksheet: {
        name: sheet.name,
        sheetId: sheetNumericId,
        dimension,
        cells,
        ...(Object.keys(formulas).length > 0 && { formulas }),
        ...(includeStyles && Object.keys(styles).length > 0 && { styles }),
        borders: {},
      },
    };
  });
}

export interface GetRangeAsCsvResult {
  success: boolean;
  csv: string;
  rowCount: number;
  columnCount: number;
  hasMore: boolean;
  sheetName: string;
}

export async function getRangeAsCsv(
  sheetId: number,
  rangeAddr: string,
  options: { includeHeaders?: boolean; maxRows?: number } = {}
): Promise<GetRangeAsCsvResult> {
  const { maxRows = 500 } = options;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    sheet.load("name");
    const range = sheet.getRange(rangeAddr);
    range.load("values,rowCount,columnCount");
    await context.sync();

    const actualRows = Math.min(range.rowCount, maxRows);
    const hasMore = range.rowCount > maxRows;

    const rows: string[] = [];
    for (let r = 0; r < actualRows; r++) {
      const row = range.values[r].map((v) => {
        if (v === null || v === undefined) return "";
        const str = String(v);
        if (str.includes(",") || str.includes('"') || str.includes("\n")) {
          return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
      });
      rows.push(row.join(","));
    }

    return {
      success: true,
      csv: rows.join("\n"),
      rowCount: actualRows,
      columnCount: range.columnCount,
      hasMore,
      sheetName: sheet.name,
    };
  });
}

export interface SearchMatch {
  sheetName: string;
  sheetId: number;
  a1: string;
  value: string | number | boolean;
  formula: string | null;
  row: number;
  column: number;
}

export interface SearchDataResult {
  success: boolean;
  matches: SearchMatch[];
  totalFound: number;
  returned: number;
  offset: number;
  hasMore: boolean;
  searchTerm: string;
  searchScope: string;
  nextOffset: number | null;
}

export async function searchData(
  searchTerm: string,
  options: {
    sheetId?: number;
    range?: string;
    offset?: number;
    matchCase?: boolean;
    matchEntireCell?: boolean;
    matchFormulas?: boolean;
    useRegex?: boolean;
    maxResults?: number;
  } = {}
): Promise<SearchDataResult> {
  const {
    sheetId,
    range,
    offset = 0,
    matchCase = false,
    matchEntireCell = false,
    matchFormulas = false,
    useRegex = false,
    maxResults = 500,
  } = options;

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    const matches: SearchMatch[] = [];
    const sheetsToSearch = sheetId
      ? [await getWorksheetById(context, sheetId)].filter(Boolean) as Excel.Worksheet[]
      : sheets.items;

    const pattern = useRegex
      ? new RegExp(searchTerm, matchCase ? "" : "i")
      : null;

    for (const sheet of sheetsToSearch) {
      sheet.load("name,id");
      const searchRange = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
      searchRange.load("values,formulas,address,rowCount,columnCount");
      await context.sync();

      if (searchRange.isNullObject) continue;

      const { startCol, startRow } = parseRangeAddress(searchRange.address);

      for (let r = 0; r < searchRange.rowCount; r++) {
        for (let c = 0; c < searchRange.columnCount; c++) {
          if (matches.length >= offset + maxResults) break;

          const value = searchRange.values[r][c];
          const formula = searchRange.formulas[r][c];
          const searchTarget = matchFormulas && formula ? String(formula) : String(value ?? "");

          let isMatch = false;
          if (pattern) {
            isMatch = pattern.test(searchTarget);
          } else {
            const compareVal = matchCase ? searchTarget : searchTarget.toLowerCase();
            const compareTerm = matchCase ? searchTerm : searchTerm.toLowerCase();
            isMatch = matchEntireCell ? compareVal === compareTerm : compareVal.includes(compareTerm);
          }

          if (isMatch && matches.length >= offset) {
            const sheetNumericId = Number.parseInt(sheet.id.replace(/\D/g, ""), 10);
            matches.push({
              sheetName: sheet.name,
              sheetId: sheetNumericId,
              a1: cellAddress(startRow + r, startCol + c),
              value: value as string | number | boolean,
              formula: typeof formula === "string" && formula.startsWith("=") ? formula : null,
              row: startRow + r + 1,
              column: startCol + c + 1,
            });
          }
        }
      }
    }

    const returned = matches.slice(0, maxResults);
    return {
      success: true,
      matches: returned,
      totalFound: matches.length + offset,
      returned: returned.length,
      offset,
      hasMore: matches.length > maxResults,
      searchTerm,
      searchScope: sheetId ? `Sheet ${sheetId}` : "All sheets",
      nextOffset: matches.length > maxResults ? offset + maxResults : null,
    };
  });
}

export interface ExcelObject {
  id: string;
  type: "chart" | "pivotTable";
  name: string;
  sheetId: number;
  sheetName: string;
}

export interface GetAllObjectsResult {
  success: boolean;
  objects: ExcelObject[];
}

export async function getAllObjects(
  options: { sheetId?: number; id?: string } = {}
): Promise<GetAllObjectsResult> {
  const { sheetId, id } = options;

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    const objects: ExcelObject[] = [];
    const sheetsToCheck = sheetId
      ? [await getWorksheetById(context, sheetId)].filter(Boolean) as Excel.Worksheet[]
      : sheets.items;

    for (const sheet of sheetsToCheck) {
      sheet.load("name,id");
      const charts = sheet.charts;
      const pivotTables = sheet.pivotTables;
      charts.load("items");
      pivotTables.load("items");
      await context.sync();

      const sheetNumericId = Number.parseInt(sheet.id.replace(/\D/g, ""), 10);

      for (const chart of charts.items) {
        chart.load("id,name");
        await context.sync();
        if (!id || chart.id === id) {
          objects.push({
            id: chart.id,
            type: "chart",
            name: chart.name,
            sheetId: sheetNumericId,
            sheetName: sheet.name,
          });
        }
      }

      for (const pivot of pivotTables.items) {
        pivot.load("id,name");
        await context.sync();
        if (!id || pivot.id === id) {
          objects.push({
            id: pivot.id,
            type: "pivotTable",
            name: pivot.name,
            sheetId: sheetNumericId,
            sheetName: sheet.name,
          });
        }
      }
    }

    return { success: true, objects };
  });
}

export interface CellInput {
  value?: unknown;
  formula?: string;
  note?: string;
  cellStyles?: {
    fontWeight?: "normal" | "bold";
    fontStyle?: "normal" | "italic";
    fontLine?: "none" | "underline" | "line-through";
    fontSize?: number;
    fontFamily?: string;
    fontColor?: string;
    backgroundColor?: string;
    horizontalAlignment?: "left" | "center" | "right";
    numberFormat?: string;
  };
  borderStyles?: {
    top?: { style?: string; weight?: string; color?: string };
    bottom?: { style?: string; weight?: string; color?: string };
    left?: { style?: string; weight?: string; color?: string };
    right?: { style?: string; weight?: string; color?: string };
  };
}

export interface SetCellRangeResult {
  success: boolean;
  cellsWritten: number;
  formulaResults?: Record<string, unknown>;
}

export async function setCellRange(
  sheetId: number,
  rangeAddr: string,
  cells: CellInput[][],
  options: {
    copyToRange?: string;
    resizeWidth?: { type: "points" | "standard"; value: number };
    resizeHeight?: { type: "points" | "standard"; value: number };
    allowOverwrite?: boolean;
  } = {}
): Promise<SetCellRangeResult> {
  const { copyToRange, resizeWidth, resizeHeight } = options;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    const range = sheet.getRange(rangeAddr);
    range.load("rowCount,columnCount");
    await context.sync();

    const values: unknown[][] = [];
    const formulas: (string | null)[][] = [];
    let hasFormulas = false;

    for (let r = 0; r < cells.length; r++) {
      values[r] = [];
      formulas[r] = [];
      for (let c = 0; c < cells[r].length; c++) {
        const cell = cells[r][c];
        if (cell.formula) {
          formulas[r][c] = cell.formula;
          values[r][c] = null;
          hasFormulas = true;
        } else {
          values[r][c] = cell.value ?? null;
          formulas[r][c] = null;
        }
      }
    }

    if (hasFormulas) {
      range.formulas = formulas.map((row, r) =>
        row.map((f, c) => f ?? values[r][c])
      );
    } else {
      range.values = values;
    }

    for (let r = 0; r < cells.length; r++) {
      for (let c = 0; c < cells[r].length; c++) {
        const cell = cells[r][c];
        if (!cell.cellStyles && !cell.note) continue;

        const cellRange = range.getCell(r, c);

        if (cell.cellStyles) {
          const s = cell.cellStyles;
          if (s.fontWeight === "bold") cellRange.format.font.bold = true;
          if (s.fontStyle === "italic") cellRange.format.font.italic = true;
          if (s.fontLine === "underline") cellRange.format.font.underline = "Single";
          if (s.fontLine === "line-through") cellRange.format.font.strikethrough = true;
          if (s.fontSize) cellRange.format.font.size = s.fontSize;
          if (s.fontFamily) cellRange.format.font.name = s.fontFamily;
          if (s.fontColor) cellRange.format.font.color = s.fontColor;
          if (s.backgroundColor) cellRange.format.fill.color = s.backgroundColor;
          if (s.horizontalAlignment) {
            cellRange.format.horizontalAlignment = s.horizontalAlignment as Excel.HorizontalAlignment;
          }
          if (s.numberFormat) cellRange.numberFormat = [[s.numberFormat]];
        }

        if (cell.note) {
          cellRange.load("address");
          await context.sync();
        }
      }
    }

    await context.sync();

    if (copyToRange) {
      const destRange = sheet.getRange(copyToRange);
      range.copyFrom(range, Excel.RangeCopyType.all);
      await context.sync();
    }

    if (resizeWidth) {
      const cols = range.getEntireColumn();
      cols.format.columnWidth = resizeWidth.value;
    }
    if (resizeHeight) {
      const rows = range.getEntireRow();
      rows.format.rowHeight = resizeHeight.value;
    }

    await context.sync();

    const formulaResults: Record<string, unknown> = {};
    if (hasFormulas) {
      range.load("values,address");
      await context.sync();
      const { startCol, startRow } = parseRangeAddress(range.address);
      for (let r = 0; r < range.values.length; r++) {
        for (let c = 0; c < range.values[r].length; c++) {
          if (formulas[r]?.[c]) {
            formulaResults[cellAddress(startRow + r, startCol + c)] = range.values[r][c];
          }
        }
      }
    }

    return {
      success: true,
      cellsWritten: cells.flat().length,
      ...(Object.keys(formulaResults).length > 0 && { formulaResults }),
    };
  });
}

export interface ClearCellRangeResult {
  success: boolean;
  clearedRange: string;
}

export async function clearCellRange(
  sheetId: number,
  rangeAddr: string,
  clearType: "contents" | "all" | "formats" = "contents"
): Promise<ClearCellRangeResult> {
  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    const range = sheet.getRange(rangeAddr);

    switch (clearType) {
      case "contents":
        range.clear(Excel.ClearApplyTo.contents);
        break;
      case "formats":
        range.clear(Excel.ClearApplyTo.formats);
        break;
      case "all":
        range.clear(Excel.ClearApplyTo.all);
        break;
    }

    await context.sync();

    return { success: true, clearedRange: rangeAddr };
  });
}

export interface CopyToResult {
  success: boolean;
  source: string;
  destination: string;
}

export async function copyTo(
  sheetId: number,
  sourceRange: string,
  destinationRange: string
): Promise<CopyToResult> {
  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    const source = sheet.getRange(sourceRange);
    const dest = sheet.getRange(destinationRange);
    dest.copyFrom(source, Excel.RangeCopyType.all);
    await context.sync();

    return { success: true, source: sourceRange, destination: destinationRange };
  });
}

export interface ModifySheetStructureResult {
  success: boolean;
  operation: string;
}

export async function modifySheetStructure(
  sheetId: number,
  params: {
    operation: "insert" | "delete" | "hide" | "unhide" | "freeze" | "unfreeze";
    dimension: "rows" | "columns";
    reference?: string;
    count?: number;
    position?: "before" | "after";
  }
): Promise<ModifySheetStructureResult> {
  const { operation, dimension, reference, count = 1, position = "before" } = params;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    if (operation === "freeze") {
      sheet.freezePanes.freezeAt(sheet.getRange(reference || "A1"));
    } else if (operation === "unfreeze") {
      sheet.freezePanes.unfreeze();
    } else if (reference) {
      const isRow = dimension === "rows";
      const rangeRef = isRow ? `${reference}:${reference}` : `${reference}:${reference}`;
      const targetRange = sheet.getRange(rangeRef);

      switch (operation) {
        case "insert":
          if (isRow) {
            targetRange.insert(position === "before" ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.down);
          } else {
            targetRange.insert(position === "before" ? Excel.InsertShiftDirection.right : Excel.InsertShiftDirection.right);
          }
          break;
        case "delete":
          targetRange.delete(isRow ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
          break;
        case "hide":
          if (isRow) {
            targetRange.rowHidden = true;
          } else {
            targetRange.columnHidden = true;
          }
          break;
        case "unhide":
          if (isRow) {
            targetRange.rowHidden = false;
          } else {
            targetRange.columnHidden = false;
          }
          break;
      }
    }

    await context.sync();
    return { success: true, operation };
  });
}

export interface ModifyWorkbookStructureResult {
  success: boolean;
  operation: string;
  sheetId?: number;
  sheetName?: string;
}

export async function modifyWorkbookStructure(params: {
  operation: "create" | "delete" | "rename" | "duplicate";
  sheetId?: number;
  sheetName?: string;
  newName?: string;
  rows?: number;
  columns?: number;
  tabColor?: string;
}): Promise<ModifyWorkbookStructureResult> {
  const { operation, sheetId, sheetName, newName, tabColor } = params;

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    switch (operation) {
      case "create": {
        const newSheet = sheets.add(sheetName);
        if (tabColor) newSheet.tabColor = tabColor;
        newSheet.load("id,name");
        await context.sync();
        const numericId = Number.parseInt(newSheet.id.replace(/\D/g, ""), 10);
        return { success: true, operation, sheetId: numericId, sheetName: newSheet.name };
      }
      case "delete": {
        const sheet = await getWorksheetById(context, sheetId!);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        sheet.delete();
        await context.sync();
        return { success: true, operation };
      }
      case "rename": {
        const sheet = await getWorksheetById(context, sheetId!);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        sheet.name = newName!;
        await context.sync();
        return { success: true, operation, sheetName: newName };
      }
      case "duplicate": {
        const sheet = await getWorksheetById(context, sheetId!);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        const copy = sheet.copy();
        if (newName) copy.name = newName;
        copy.load("id,name");
        await context.sync();
        const numericId = Number.parseInt(copy.id.replace(/\D/g, ""), 10);
        return { success: true, operation, sheetId: numericId, sheetName: copy.name };
      }
    }
  });
}

export interface ResizeRangeResult {
  success: boolean;
}

export async function resizeRange(
  sheetId: number,
  params: {
    range?: string;
    width?: { type: "points" | "standard"; value: number };
    height?: { type: "points" | "standard"; value: number };
  }
): Promise<ResizeRangeResult> {
  const { range, width, height } = params;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    const targetRange = range ? sheet.getRange(range) : sheet.getUsedRange();

    if (width) {
      const cols = targetRange.getEntireColumn();
      cols.format.columnWidth = width.value;
    }
    if (height) {
      const rows = targetRange.getEntireRow();
      rows.format.rowHeight = height.value;
    }

    await context.sync();
    return { success: true };
  });
}

export interface ModifyObjectResult {
  success: boolean;
  operation: string;
  id?: string;
}

export interface SheetMetadata {
  id: number;
  name: string;
  maxRows: number;
  maxColumns: number;
  frozenRows: number;
  frozenColumns: number;
}

export interface WorkbookMetadata {
  success: boolean;
  fileName: string;
  sheetsMetadata: SheetMetadata[];
  totalSheets: number;
}

export async function getWorkbookMetadata(): Promise<WorkbookMetadata> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    workbook.load("name");
    const sheets = workbook.worksheets;
    sheets.load("items");
    await context.sync();

    const sheetData: {
      sheet: Excel.Worksheet;
      usedRange: Excel.Range;
      freezeLocation: Excel.Range;
    }[] = [];

    for (const sheet of sheets.items) {
      sheet.load("id,name");
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("rowCount,columnCount");
      const freezeLocation = sheet.freezePanes.getLocationOrNullObject();
      freezeLocation.load("rowCount,columnCount");
      sheetData.push({ sheet, usedRange, freezeLocation });
    }
    await context.sync();

    const sheetsMetadata: SheetMetadata[] = sheetData.map(({ sheet, usedRange, freezeLocation }) => ({
      id: Number.parseInt(sheet.id.replace(/\D/g, ""), 10),
      name: sheet.name,
      maxRows: usedRange.isNullObject ? 0 : usedRange.rowCount,
      maxColumns: usedRange.isNullObject ? 0 : usedRange.columnCount,
      frozenRows: freezeLocation.isNullObject ? 0 : freezeLocation.rowCount,
      frozenColumns: freezeLocation.isNullObject ? 0 : freezeLocation.columnCount,
    }));

    return {
      success: true,
      fileName: workbook.name || "Untitled",
      sheetsMetadata,
      totalSheets: sheets.items.length,
    };
  });
}

export async function modifyObject(params: {
  operation: "create" | "update" | "delete";
  sheetId: number;
  objectType: "pivotTable" | "chart";
  id?: string;
  properties?: {
    name?: string;
    source?: string;
    range?: string;
    anchor?: string;
    rows?: { field: string }[];
    columns?: { field: string }[];
    values?: { field: string; summarizeBy?: string }[];
    title?: string;
    chartType?: string;
  };
}): Promise<ModifyObjectResult> {
  const { operation, sheetId, objectType, id, properties } = params;

  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    if (objectType === "chart") {
      const charts = sheet.charts;

      switch (operation) {
        case "create": {
          if (!properties?.source || !properties?.chartType) {
            throw new Error("Chart creation requires source and chartType");
          }
          const sourceRange = sheet.getRange(properties.source);
          const chart = charts.add(
            properties.chartType as Excel.ChartType,
            sourceRange,
            Excel.ChartSeriesBy.auto
          );
          if (properties.title) chart.title.text = properties.title;
          if (properties.anchor) {
            const anchorCell = sheet.getRange(properties.anchor);
            chart.setPosition(anchorCell);
          }
          chart.load("id");
          await context.sync();
          return { success: true, operation, id: chart.id };
        }
        case "update": {
          if (!id) throw new Error("Chart update requires id");
          const chart = charts.getItem(id);
          if (properties?.title) chart.title.text = properties.title;
          await context.sync();
          return { success: true, operation, id };
        }
        case "delete": {
          if (!id) throw new Error("Chart delete requires id");
          const chart = charts.getItem(id);
          chart.delete();
          await context.sync();
          return { success: true, operation };
        }
      }
    } else {
      const pivotTables = sheet.pivotTables;

      switch (operation) {
        case "create": {
          if (!properties?.source || !properties?.range) {
            throw new Error("PivotTable creation requires source and range");
          }
          const sourceRange = sheet.getRange(properties.source);
          const destRange = sheet.getRange(properties.range);
          const pivot = context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
            properties.name || "PivotTable",
            sourceRange,
            destRange
          );
          pivot.load("id");
          await context.sync();
          return { success: true, operation, id: pivot.id };
        }
        case "update": {
          if (!id) throw new Error("PivotTable update requires id");
          const pivot = pivotTables.getItem(id);
          if (properties?.name) pivot.name = properties.name;
          await context.sync();
          return { success: true, operation, id };
        }
        case "delete": {
          if (!id) throw new Error("PivotTable delete requires id");
          const pivot = pivotTables.getItem(id);
          pivot.delete();
          await context.sync();
          return { success: true, operation };
        }
      }
    }

    return { success: false, operation: "unknown" };
  });
}
