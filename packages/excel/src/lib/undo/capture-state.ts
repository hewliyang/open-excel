/**
 * State capture utilities for undo functionality
 *
 * These functions capture the current state of cells/ranges before
 * operations are performed, allowing us to restore them later.
 */

export interface CellState {
  address: string;
  values: any[][];
  formulas: string[][];
  numberFormats: string[][];
  formats?: {
    fill?: any;
    font?: any;
    borders?: any;
  }[][];
}

export interface TableState {
  name: string;
  address: string;
  range: string;
  columns: Array<{ name: string; index: number }>;
  data: any[][];
  headerRowCount: number;
  showTotals: boolean;
}

export interface SheetState {
  name: string;
  position: number;
  visibility: Excel.SheetVisibility;
}

/**
 * Capture current state of a cell range
 */
export async function captureCellRangeState(
  sheetName: string,
  address: string
): Promise<CellState> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load([
      "address",
      "values",
      "formulas",
      "numberFormat",
      "format/fill",
      "format/font",
      "format/borders",
    ]);

    await context.sync();

    return {
      address: range.address,
      values: range.values,
      formulas: range.formulas,
      numberFormats: range.numberFormat,
      formats: [[range.format]], // Simplified for now
    };
  });
}

/**
 * Restore cell range state
 */
export async function restoreCellRangeState(
  sheetName: string,
  state: CellState
): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(state.address);

    // Restore formulas (which includes values for non-formula cells)
    range.formulas = state.formulas;
    range.numberFormat = state.numberFormats;

    await context.sync();
  });
}

/**
 * Capture current state of a table
 */
export async function captureTableState(
  sheetName: string,
  tableName: string
): Promise<TableState> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItem(tableName);

    table.load([
      "name",
      "range/address",
      "columns",
      "headerRowCount",
      "showTotals",
    ]);

    // Load all data including headers
    const dataRange = table.getRange();
    dataRange.load(["values"]);

    await context.sync();

    const columns = table.columns.items.map((col, index) => ({
      name: col.name,
      index,
    }));

    return {
      name: table.name,
      address: dataRange.address,
      range: dataRange.address,
      columns,
      data: dataRange.values,
      headerRowCount: table.headerRowCount,
      showTotals: table.showTotals,
    };
  });
}

/**
 * Restore table state
 */
export async function restoreTableState(
  sheetName: string,
  state: TableState
): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);

    // Check if table exists
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    const existingTable = tables.items.find((t) => t.name === state.name);

    if (existingTable) {
      // Table exists, restore its data
      existingTable.load("range");
      await context.sync();

      const tableRange = existingTable.getRange();
      tableRange.values = state.data;
    } else {
      // Table doesn't exist, recreate it
      const range = sheet.getRange(state.range);
      range.values = state.data;

      const newTable = sheet.tables.add(state.range, true);
      newTable.name = state.name;
      newTable.showTotals = state.showTotals;

      // Rename columns
      newTable.load("columns");
      await context.sync();

      state.columns.forEach((col, index) => {
        if (newTable.columns.items[index]) {
          newTable.columns.items[index].name = col.name;
        }
      });
    }

    await context.sync();
  });
}

/**
 * Capture sheet state
 */
export async function captureSheetState(
  sheetName: string
): Promise<SheetState> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load(["name", "position", "visibility"]);
    await context.sync();

    return {
      name: sheet.name,
      position: sheet.position,
      visibility: sheet.visibility,
    };
  });
}

/**
 * Delete a sheet (for undo of sheet creation)
 */
export async function deleteSheet(sheetName: string): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.delete();
    await context.sync();
  });
}

/**
 * Recreate a sheet (for undo of sheet deletion)
 */
export async function recreateSheet(state: SheetState): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add(state.name);
    sheet.position = state.position;
    sheet.visibility = state.visibility;
    await context.sync();
  });
}
