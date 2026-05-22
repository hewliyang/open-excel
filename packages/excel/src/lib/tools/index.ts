export { createBashTool, createReadTool } from "@office-agents/core";
export { clearCellRangeTool } from "./clear-cell-range";
export { copyToTool } from "./copy-to";
export { createEvalOfficeJsTool } from "./eval-officejs";
export { getAllObjectsTool } from "./get-all-objects";
export { getCellRangesTool } from "./get-cell-ranges";
export { getRangeAsCsvTool } from "./get-range-as-csv";
export { modifyObjectTool } from "./modify-object";
export { modifySheetStructureTool } from "./modify-sheet-structure";
export { modifyWorkbookStructureTool } from "./modify-workbook-structure";
export { resizeRangeTool } from "./resize-range";
export { screenshotRangeTool } from "./screenshot-range";
export { searchDataTool } from "./search-data";
export { setCellRangeTool } from "./set-cell-range";
export { undoTool, undoHistoryTool } from "./undo";
export { checkPermissionsTool } from "./check-permissions";
export { provideFormulasTool } from "./provide-formulas";
export {
  defineTool,
  type ToolResult,
  toolError,
  toolSuccess,
  toolText,
} from "./types";

import type { AgentContext } from "@office-agents/core";
import { createBashTool, createReadTool } from "@office-agents/core";
import { clearCellRangeTool } from "./clear-cell-range";
import { copyToTool } from "./copy-to";
import { createEvalOfficeJsTool } from "./eval-officejs";
import { getAllObjectsTool } from "./get-all-objects";
import { getCellRangesTool } from "./get-cell-ranges";
import { getRangeAsCsvTool } from "./get-range-as-csv";
import { modifyObjectTool } from "./modify-object";
import { modifySheetStructureTool } from "./modify-sheet-structure";
import { modifyWorkbookStructureTool } from "./modify-workbook-structure";
import { resizeRangeTool } from "./resize-range";
import { screenshotRangeTool } from "./screenshot-range";
import { searchDataTool } from "./search-data";
import { setCellRangeTool } from "./set-cell-range";
import { undoHistoryTool, undoTool } from "./undo";
import { checkPermissionsTool } from "./check-permissions";
import { provideFormulasTool } from "./provide-formulas";

export function createExcelTools(ctx: AgentContext) {
  return [
    // fs tools
    createReadTool(ctx),
    createBashTool(ctx),
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
    createEvalOfficeJsTool(ctx),
    // Undo/redo
    undoTool,
    undoHistoryTool,
    // Safety & fallback tools
    checkPermissionsTool,
    provideFormulasTool,
  ];
}
