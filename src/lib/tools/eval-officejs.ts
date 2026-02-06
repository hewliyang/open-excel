import { Type } from "@sinclair/typebox";
import { ensureLockdown } from "../../taskpane/lockdown";
import type { DirtyRange } from "../dirty-tracker";
import { createTrackedContext } from "../excel/tracked-context";
import { defineTool, toolError, toolSuccess } from "./types";

/* global Excel, Compartment */

const MUTATION_PATTERNS = [
  /\.(values|formulas|numberFormat)\s*=/,
  /\.clear\s*\(/,
  /\.delete\s*\(/,
  /\.insert\s*\(/,
  /\.copyFrom\s*\(/,
  /\.add\s*\(/,
];

function looksLikeMutation(code: string): boolean {
  return MUTATION_PATTERNS.some((p) => p.test(code));
}

const BLOCKED_OBJECT_METHODS = new Set([
  "defineProperty",
  "getOwnPropertyDescriptor",
  "getPrototypeOf",
  "setPrototypeOf",
]);

function createRestrictedObject(): Record<string, unknown> {
  const restricted: Record<string, unknown> = {};
  for (const key of Object.getOwnPropertyNames(Object)) {
    if (!BLOCKED_OBJECT_METHODS.has(key)) {
      restricted[key] = (Object as unknown as Record<string, unknown>)[key];
    }
  }
  return restricted;
}

function sandboxedEval(code: string, globals: Record<string, unknown>): unknown {
  ensureLockdown();
  const compartment = new Compartment({
    globals: {
      ...globals,
      console,
      Math,
      Date,
      Object: createRestrictedObject(),
      Function: undefined,
      Reflect: undefined,
      Proxy: undefined,
      Compartment: undefined,
      harden: undefined,
      lockdown: undefined,
    },
    __options__: true, // required to use options-bag constructor form
  });
  return compartment.evaluate(`(async () => { ${code} })()`);
}

export const evalOfficeJsTool = defineTool({
  name: "eval_officejs",
  label: "Execute Office.js Code",
  description:
    "Execute arbitrary Office.js code within an Excel.run context. " +
    "Use this as an escape hatch when existing tools don't cover your use case. " +
    "The code runs inside `Excel.run(async (context) => { ... })` with `context` available. " +
    "Return a value to get it back as the result. Always call `await context.sync()` before returning.",
  parameters: Type.Object({
    code: Type.String({
      description:
        "JavaScript code to execute. Has access to `context` (Excel.RequestContext). " +
        "Must be valid async code. Return a value to get it as result. " +
        "Example: `const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1'); range.load('values'); await context.sync(); return range.values;`",
    }),
    explanation: Type.Optional(
      Type.String({
        description: "Brief explanation of what this code does (max 100 chars)",
        maxLength: 100,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      let dirtyRanges: DirtyRange[] = [];

      const result = await Excel.run(async (context) => {
        const { trackedContext, getDirtyRanges } = createTrackedContext(context);

        const execResult = await sandboxedEval(params.code, {
          context: trackedContext,
          Excel,
        });

        dirtyRanges = getDirtyRanges();
        return execResult;
      });

      if (dirtyRanges.length === 0 && looksLikeMutation(params.code)) {
        dirtyRanges = [{ sheetId: -1, range: "*" }];
      }

      const response: Record<string, unknown> = { success: true, result: result ?? null };
      if (dirtyRanges.length > 0) {
        response._dirtyRanges = dirtyRanges;
      }
      return toolSuccess(response);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error executing code";
      return toolError(message);
    }
  },
});
