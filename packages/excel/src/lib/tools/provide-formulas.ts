import { Type } from "@sinclair/typebox";
import { defineTool, toolSuccess } from "./types";

/**
 * Provide copy-paste ready formulas
 *
 * When direct writes are blocked, generate formulas that users can manually paste.
 * This is a fallback when Excel protection prevents programmatic writes.
 */
export const provideFormulasTool = defineTool({
  name: "provide_copy_paste_formulas",
  description:
    "Generate copy-paste ready formulas for users to manually enter when direct writes are blocked. Use this as a fallback when set_cell_range fails due to protection.",
  parameters: Type.Object({
    sheetName: Type.String({ description: "Target sheet name" }),
    startCell: Type.String({ description: "Starting cell (e.g., A1)" }),
    layout: Type.String({
      description:
        "Layout description (e.g., 'Row 1 headers, Row 2 ASP Price formulas')",
    }),
    formulas: Type.Array(
      Type.Object({
        cell: Type.String({ description: "Cell address (e.g., A1)" }),
        formula: Type.String({
          description: "Formula or value (e.g., =HighLevel!G28)",
        }),
        description: Type.Optional(
          Type.String({ description: "What this cell calculates" })
        ),
      }),
      {
        description: "Array of formulas to provide",
      }
    ),
  }),
  execute: async (_toolCallId, params) => {
    let output = `📋 **Copy-Paste Ready Formulas for "${params.sheetName}"**\n\n`;
    output += `Since Excel is blocking direct writes, please manually paste these formulas:\n\n`;
    output += `**Layout:** ${params.layout}\n`;
    output += `**Starting at:** ${params.startCell}\n\n`;
    output += `---\n\n`;

    // Group formulas by row for easier copying
    const formulasByRow: { [row: string]: typeof params.formulas } = {};

    params.formulas.forEach((f) => {
      const match = f.cell.match(/^([A-Z]+)(\d+)$/);
      if (match) {
        const row = match[2];
        if (!formulasByRow[row]) {
          formulasByRow[row] = [];
        }
        formulasByRow[row].push(f);
      }
    });

    // Output formulas row by row
    Object.keys(formulasByRow)
      .sort((a, b) => parseInt(a) - parseInt(b))
      .forEach((row) => {
        output += `**Row ${row}:**\n`;
        formulasByRow[row].forEach((f) => {
          output += `- **${f.cell}**: \`${f.formula}\``;
          if (f.description) {
            output += ` (${f.description})`;
          }
          output += "\n";
        });
        output += "\n";
      });

    output += `---\n\n`;
    output += `**How to paste:**\n`;
    output += `1. Select cell ${params.startCell} on sheet "${params.sheetName}"\n`;
    output += `2. Copy each formula above\n`;
    output += `3. Paste into the corresponding cell\n`;
    output += `4. Press Enter to confirm\n\n`;
    output += `💡 **Tip:** You can paste multiple cells at once by selecting a range first.`;

    return toolSuccess(output);
  },
});
