export { bashTool } from "./bash";
export { clearCellRangeTool } from "./clear-cell-range";
export { copyToTool } from "./copy-to";
export { evalOfficeJsTool } from "./eval-officejs";
export { getAllObjectsTool } from "./get-all-objects";
export { getCellRangesTool } from "./get-cell-ranges";
export { getRangeAsCsvTool } from "./get-range-as-csv";
export { modifyObjectTool } from "./modify-object";
export { modifySheetStructureTool } from "./modify-sheet-structure";
export { modifyWorkbookStructureTool } from "./modify-workbook-structure";
export { readTool } from "./read-file";
export { resizeRangeTool } from "./resize-range";
export { screenshotRangeTool } from "./screenshot-range";
export { searchDataTool } from "./search-data";
export { setCellRangeTool } from "./set-cell-range";
export {
  defineTool,
  type ToolResult,
  toolError,
  toolSuccess,
  toolText,
} from "./types";

import { bashTool } from "./bash";
import { clearCellRangeTool } from "./clear-cell-range";
import { copyToTool } from "./copy-to";
import { evalOfficeJsTool } from "./eval-officejs";
import { getAllObjectsTool } from "./get-all-objects";
import { getCellRangesTool } from "./get-cell-ranges";
import { getRangeAsCsvTool } from "./get-range-as-csv";
import { modifyObjectTool } from "./modify-object";
import { modifySheetStructureTool } from "./modify-sheet-structure";
import { modifyWorkbookStructureTool } from "./modify-workbook-structure";
import { readTool } from "./read-file";
import { resizeRangeTool } from "./resize-range";
import { screenshotRangeTool } from "./screenshot-range";
import { searchDataTool } from "./search-data";
import { setCellRangeTool } from "./set-cell-range";

export const EXCEL_TOOLS = [
  // fs tools
  readTool,
  bashTool,
  // Excel read tools
  getCellRangesTool,
  getRangeAsCsvTool,
  searchDataTool,
  screenshotRangeTool,
  getAllObjectsTool,
  // Excel write tools
  setCellRangeTool,
  clearCellRangeTool,
  copyToTool,
  modifySheetStructureTool,
  modifyWorkbookStructureTool,
  resizeRangeTool,
  modifyObjectTool,
  evalOfficeJsTool,
];
