import { Type } from "@sinclair/typebox";
import { setCellRange } from "../excel/api";
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
      return toolSuccess(result);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unknown error writing cells";
      return toolError(message);
    }
  },
});
