import { Type } from "@sinclair/typebox";
import type { AgentContext } from "../context";
import { defineTool, toolError, toolSuccess } from "./types";

interface EditOp {
  old_text: string;
  new_text: string;
}

function resolvePath(path: string): string {
  return path.startsWith("/") ? path : `/home/user/uploads/${path}`;
}

function countOccurrences(haystack: string, needle: string): number {
  if (needle.length === 0) return 0;
  let count = 0;
  let idx = 0;
  while (true) {
    const found = haystack.indexOf(needle, idx);
    if (found === -1) break;
    count++;
    idx = found + needle.length;
  }
  return count;
}

export function createEditFileTool(ctx: AgentContext) {
  return defineTool({
    name: "edit_file",
    label: "Edit File",
    description:
      "Write or edit a file in the virtual filesystem. " +
      "Pass `content` to create or overwrite the entire file. " +
      "Pass `edits` (array of { old_text, new_text }) to apply targeted replacements; " +
      "each old_text must match exactly once in the current file contents. " +
      "Exactly one of `content` or `edits` must be provided. " +
      "Parent directories are created automatically.",
    parameters: Type.Object({
      path: Type.String({
        description:
          "File path. Absolute (starting with /) or relative to /home/user/uploads/.",
      }),
      content: Type.Optional(
        Type.String({
          description:
            "Full file contents. Creates the file if missing, overwrites if present. " +
            "Mutually exclusive with `edits`.",
        }),
      ),
      edits: Type.Optional(
        Type.Array(
          Type.Object({
            old_text: Type.String({
              description:
                "Exact text to find. Must match exactly once in the file.",
            }),
            new_text: Type.String({
              description: "Replacement text.",
            }),
          }),
          {
            description:
              "Ordered list of replacements applied sequentially. Each old_text " +
              "must be unique in the file at the moment it is applied.",
          },
        ),
      ),
      explanation: Type.Optional(
        Type.String({
          description: "Brief explanation (max 50 chars)",
          maxLength: 50,
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      const hasContent = params.content !== undefined;
      const hasEdits = Array.isArray(params.edits) && params.edits.length > 0;

      if (hasContent && hasEdits) {
        return toolError(
          "Pass either `content` (full rewrite) or `edits` (targeted replacements), not both.",
        );
      }
      if (!hasContent && !hasEdits) {
        return toolError(
          "Must provide either `content` (full rewrite) or a non-empty `edits` array.",
        );
      }

      const fullPath = resolvePath(params.path);

      try {
        if (hasContent) {
          const existed = await ctx.fileExists(fullPath);
          await ctx.writeFile(fullPath, params.content as string);
          const bytes = new TextEncoder().encode(
            params.content as string,
          ).length;
          return toolSuccess({
            success: true,
            path: fullPath,
            action: existed ? "overwrote" : "created",
            bytes,
          });
        }

        if (!(await ctx.fileExists(fullPath))) {
          return toolError(
            `File not found: ${fullPath}. Use \`content\` to create a new file, or write it first.`,
          );
        }

        let text = await ctx.readFile(fullPath);
        const applied: Array<{ index: number; bytesDelta: number }> = [];

        const edits = params.edits as EditOp[];
        for (let i = 0; i < edits.length; i++) {
          const { old_text, new_text } = edits[i];
          if (old_text.length === 0) {
            return toolError(`Edit #${i + 1}: old_text must not be empty.`);
          }
          const occurrences = countOccurrences(text, old_text);
          if (occurrences === 0) {
            return toolError(
              `Edit #${i + 1}: old_text not found in ${fullPath}. ` +
                "It must match exactly (including whitespace).",
            );
          }
          if (occurrences > 1) {
            return toolError(
              `Edit #${i + 1}: old_text matches ${occurrences} times in ${fullPath}. ` +
                "Provide more surrounding context so it is unique.",
            );
          }
          const before = text.length;
          text = text.replace(old_text, new_text);
          applied.push({ index: i + 1, bytesDelta: text.length - before });
        }

        await ctx.writeFile(fullPath, text);
        return toolSuccess({
          success: true,
          path: fullPath,
          action: "edited",
          edits: applied.length,
          bytes: new TextEncoder().encode(text).length,
        });
      } catch (error) {
        const msg =
          error instanceof Error ? error.message : "Failed to edit file";
        return toolError(msg);
      }
    },
  });
}
