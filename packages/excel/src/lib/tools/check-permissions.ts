import { defineTool, toolSuccess, toolError } from "./types";

/**
 * Check workbook write permissions
 *
 * Before performing write operations, check if the workbook allows writes.
 * If blocked, guide the user to manually enable write access.
 */
export const checkPermissionsTool = defineTool({
  name: "check_write_permissions",
  description:
    "Check if the workbook allows write operations. Use this before attempting writes if you suspect permission issues.",
  parameters: {
    type: "object",
    properties: {},
  },
  execute: async () => {
    try {
      const status = await Excel.run(async (context) => {
        const workbook = context.workbook;
        const protection = workbook.protection;

        protection.load("protected");
        await context.sync();

        const sheets = context.workbook.worksheets;
        sheets.load("items/protection/protected");
        await context.sync();

        const protectedSheets = sheets.items
          .filter((sheet) => sheet.protection.protected)
          .map((sheet) => sheet.name);

        return {
          workbookProtected: protection.protected,
          protectedSheets,
          canWrite: !protection.protected && protectedSheets.length === 0,
        };
      });

      if (status.canWrite) {
        return toolSuccess("✅ Workbook is writable. All write operations should work.");
      } else {
        let message = "⚠️ **Write restrictions detected:**\n\n";

        if (status.workbookProtected) {
          message += "- Workbook is protected\n";
        }

        if (status.protectedSheets.length > 0) {
          message += `- Protected sheets: ${status.protectedSheets.join(", ")}\n`;
        }

        message += "\n**To enable writes:**\n";
        message += "1. Go to Review → Unprotect Workbook/Sheet\n";
        message += "2. Or ask the user to enable write access\n\n";
        message += "I can provide copy-paste ready formulas instead.";

        return toolSuccess(message);
      }
    } catch (error) {
      return toolError(
        `Failed to check permissions: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  },
});
