import { Type } from "@sinclair/typebox";
import { setCellRange, getWorksheetById } from "../excel/api";
import {
  captureCellRangeState,
  restoreCellRangeState,
  undoManager,
} from "../undo";
import { defineTool, toolError, toolSuccess } from "./types";

const BorderStyleSchema = Type.Optional(
  Type.Object({
    style: Type.Optional(
      Type.Union([
        Type.Literal("solid"),
        Type.Literal("dashed"),
        Type.Literal("dotted"),
        Type.Literal("double"),
      ]),
    ),
    weight: Type.Optional(
      Type.Union([
        Type.Literal("thin"),
        Type.Literal("medium"),
        Type.Literal("thick"),
      ]),
    ),
    color: Type.Optional(Type.String()),
  }),
);

const CellStylesSchema = Type.Optional(
  Type.Object({
    fontWeight: Type.Optional(
      Type.Union([Type.Literal("normal"), Type.Literal("bold")]),
    ),
    fontStyle: Type.Optional(
      Type.Union([Type.Literal("normal"), Type.Literal("italic")]),
    ),
    fontLine: Type.Optional(
      Type.Union([
        Type.Literal("none"),
        Type.Literal("underline"),
        Type.Literal("line-through"),
      ]),
    ),
    fontSize: Type.Optional(Type.Number()),
    fontFamily: Type.Optional(Type.String()),
    fontColor: Type.Optional(Type.String()),
    backgroundColor: Type.Optional(Type.String()),
    horizontalAlignment: Type.Optional(
      Type.Union([
        Type.Literal("left"),
        Type.Literal("center"),
        Type.Literal("right"),
      ]),
    ),
    numberFormat: Type.Optional(Type.String()),
  }),
);

const BorderStylesSchema = Type.Optional(
  Type.Object({
    top: BorderStyleSchema,
    bottom: BorderStyleSchema,
    left: BorderStyleSchema,
    right: BorderStyleSchema,
  }),
);

const CellSchema = Type.Object({
  value: Type.Optional(Type.Any()),
  formula: Type.Optional(Type.String()),
  note: Type.Optional(Type.String()),
  cellStyles: CellStylesSchema,
  borderStyles: BorderStylesSchema,
});

const ResizeSchema = Type.Optional(
  Type.Object({
    type: Type.Union([Type.Literal("points"), Type.Literal("standard")]),
    value: Type.Number(),
  }),
);

export const setCellRangeTool = defineTool({
  name: "set_cell_range",
  label: "Set Cell Range",
  description:
    "WRITE. Write values, formulas, and formatting to cells. " +
    "The range is auto-expanded to match the cells array dimensions (e.g. A1 with a 1x3 array becomes A1:C1). " +
    "OVERWRITE PROTECTION: By default, fails if target cells contain data. " +
    "If the tool returns an overwrite error, read those cells to see what's there, " +
    "confirm with the user, then retry with allow_overwrite=true. " +
    "Only set allow_overwrite=true on first attempt if user explicitly says 'replace' or 'overwrite'. " +
    "Use copyToRange to expand a pattern to a larger area.",
  parameters: Type.Object({
    sheetId: Type.Number({ description: "The worksheet ID (1-based index)" }),
    range: Type.String({
      description:
        "Target range in A1 notation (auto-expands to match cells dimensions)",
    }),
    cells: Type.Array(Type.Array(CellSchema), {
      description: "2D array of cell data matching range dimensions",
    }),
    copyToRange: Type.Optional(
      Type.String({
        description: "Expand pattern to larger range after writing",
      }),
    ),
    resizeWidth: ResizeSchema,
    resizeHeight: ResizeSchema,
    allow_overwrite: Type.Optional(
      Type.Boolean({ description: "Confirm overwriting existing data" }),
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief explanation (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  dirtyTracking: {
    getRanges: (p) => {
      const ranges = [{ sheetId: p.sheetId, range: p.range }];
      if (p.copyToRange) {
        ranges.push({ sheetId: p.sheetId, range: p.copyToRange });
      }
      return ranges;
    },
  },
  execute: async (_toolCallId, params) => {
    try {
      // Execute the operation first
      const result = await setCellRange(
        params.sheetId,
        params.range,
        params.cells,
        {
          copyToRange: params.copyToRange,
          resizeWidth: params.resizeWidth,
          resizeHeight: params.resizeHeight,
          allowOverwrite: params.allow_overwrite,
        },
      );

      // Try to capture state for undo (best effort - don't fail if this doesn't work)
      try {
        await Excel.run(async (context) => {
          const sheet = await getWorksheetById(context, params.sheetId);
          if (!sheet) {
            console.log("[set_cell_range] Sheet not found for undo, skipping");
            return;
          }

          sheet.load("name");
          await context.sync();
          const sheetName = sheet.name;

          // Try to capture current state for potential undo
          try {
            const currentState = await captureCellRangeState(sheetName, params.range);

            // Register undo operation
            undoManager.registerOperation(
              `Set cells in ${sheetName}!${params.range}`,
              async () => {
                await restoreCellRangeState(sheetName, currentState);
              }
            );
          } catch (stateErr) {
            console.log("[set_cell_range] Could not capture state for undo:", stateErr);
          }
        });
      } catch (undoErr) {
        // Undo registration failed - that's ok, the write still succeeded
        console.log("[set_cell_range] Could not register undo:", undoErr);
      }

      return toolSuccess(result);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);

      // Detect permission/protection errors
      if (
        errorMessage.includes("protected") ||
        errorMessage.includes("permission") ||
        errorMessage.includes("read-only") ||
        errorMessage.includes("restricted")
      ) {
        return toolError(
          `❌ **Write blocked - Workbook is protected**\n\n` +
          `Cannot write to cells because Excel is blocking modifications.\n\n` +
          `**Solution:** Provide copy-paste ready formulas instead:\n` +
          `1. Use check_write_permissions tool to diagnose\n` +
          `2. Show the user the exact formulas to paste\n` +
          `3. Guide them to paste manually in the target range\n\n` +
          `Error: ${errorMessage}`
        );
      }

      return toolError(`Failed to write cells: ${errorMessage}`);
    }
  },
});
