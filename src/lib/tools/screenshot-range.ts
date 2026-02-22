/* global Excel */

import { Type } from "@sinclair/typebox";
import { parseRange } from "../dirty-tracker";
import { getWorksheetById } from "../excel/api";
import type { ToolResult } from "./types";
import { defineTool, toolError } from "./types";

const HEADER_WIDTH = 40;
const HEADER_HEIGHT = 20;
const HEADER_BG = "#f0f0f0";
const HEADER_BORDER = "#c0c0c0";
const HEADER_FONT = "bold 11px Calibri, Arial, sans-serif";
const HEADER_TEXT_COLOR = "#333333";

function columnIndexToLetter(index: number): string {
  let letter = "";
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function compositeWithHeaders(
  imageBase64: string,
  startRow: number,
  startCol: number,
  colWidths: number[],
  rowHeights: number[],
): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const totalColWidth = colWidths.reduce((a, b) => a + b, 0);
      const totalRowHeight = rowHeights.reduce((a, b) => a + b, 0);
      const scaleX = totalColWidth > 0 ? img.width / totalColWidth : 1;
      const scaleY = totalRowHeight > 0 ? img.height / totalRowHeight : 1;

      const canvas = document.createElement("canvas");
      canvas.width = HEADER_WIDTH + img.width;
      canvas.height = HEADER_HEIGHT + img.height;

      const ctx = canvas.getContext("2d");
      if (!ctx) {
        return reject(new Error("Failed to get 2d canvas context"));
      }

      // White background
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Draw the range image
      ctx.drawImage(img, HEADER_WIDTH, HEADER_HEIGHT);

      // Column headers
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(HEADER_WIDTH, 0, img.width, HEADER_HEIGHT);
      ctx.strokeStyle = HEADER_BORDER;
      ctx.font = HEADER_FONT;
      ctx.fillStyle = HEADER_TEXT_COLOR;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      let x = HEADER_WIDTH;
      for (let i = 0; i < colWidths.length; i++) {
        const w = colWidths[i] * scaleX;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.beginPath();
        ctx.moveTo(x, 0);
        ctx.lineTo(x, HEADER_HEIGHT);
        ctx.stroke();
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(
          columnIndexToLetter(startCol + i),
          x + w / 2,
          HEADER_HEIGHT / 2,
        );
        x += w;
      }

      // Row headers
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, HEADER_HEIGHT, HEADER_WIDTH, img.height);
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      let y = HEADER_HEIGHT;
      for (let i = 0; i < rowHeights.length; i++) {
        const h = rowHeights[i] * scaleY;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(HEADER_WIDTH, y);
        ctx.stroke();
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(String(startRow + i + 1), HEADER_WIDTH / 2, y + h / 2);
        y += h;
      }

      // Corner cell
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
      ctx.strokeStyle = HEADER_BORDER;
      ctx.strokeRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
      ctx.strokeRect(HEADER_WIDTH, 0, img.width, HEADER_HEIGHT);
      ctx.strokeRect(0, HEADER_HEIGHT, HEADER_WIDTH, img.height);

      resolve(canvas.toDataURL("image/png").split(",")[1]);
    };
    img.onerror = () => reject(new Error("Failed to load range image"));
    img.src = `data:image/png;base64,${imageBase64}`;
  });
}

export const screenshotRangeTool = defineTool({
  name: "screenshot_range",
  label: "Screenshot Range",
  description:
    "Capture a screenshot of a cell range as an image with row/column headers. " +
    "Returns the image for visual inspection — useful for verifying formatting, " +
    "charts embedded in cells, conditional formatting, and overall layout. " +
    "The range should not be excessively large; keep it reasonable (e.g. A1:Z50).",
  parameters: Type.Object({
    sheetId: Type.Number({ description: "The worksheet ID (1-based index)" }),
    range: Type.String({
      description: "Range in A1 notation, e.g. 'A1:F20' or 'B3:M30'",
    }),
    explanation: Type.Optional(
      Type.String({
        description:
          "Brief explanation of what you're inspecting (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params): Promise<ToolResult> => {
    try {
      const parsed = parseRange(params.range);
      if (!parsed) {
        return toolError(`Invalid range: ${params.range}`);
      }

      const { startRow, startCol, endCol, endRow } = parsed;
      const numCols = endCol - startCol + 1;
      const numRows = endRow - startRow + 1;

      const data = await Excel.run(async (context) => {
        const sheet = await getWorksheetById(context, params.sheetId);
        if (!sheet) {
          throw new Error(`Worksheet with ID ${params.sheetId} not found`);
        }

        const range = sheet.getRange(params.range);
        const image = range.getImage();

        const cols: Excel.Range[] = [];
        for (let i = 0; i < numCols; i++) {
          const col = range.getColumn(i);
          col.format.load("columnWidth");
          cols.push(col);
        }

        const rows: Excel.Range[] = [];
        for (let i = 0; i < numRows; i++) {
          const row = range.getRow(i);
          row.format.load("rowHeight");
          rows.push(row);
        }

        await context.sync();

        return {
          imageBase64: image.value,
          colWidths: cols.map((c) => c.format.columnWidth),
          rowHeights: rows.map((r) => r.format.rowHeight),
        };
      });

      const base64 = await compositeWithHeaders(
        data.imageBase64,
        startRow,
        startCol,
        data.colWidths,
        data.rowHeights,
      );

      return {
        content: [
          {
            type: "text" as const,
            text: `Screenshot of ${params.range}`,
          },
          {
            type: "image" as const,
            data: base64,
            mimeType: "image/png" as const,
          },
        ],
        details: undefined,
      };
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "Unknown error capturing screenshot";
      return toolError(message);
    }
  },
});
