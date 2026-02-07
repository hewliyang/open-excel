import { Type } from "@sinclair/typebox";
import { DEFAULT_MAX_BYTES, DEFAULT_MAX_LINES, formatSize, truncateTail } from "../truncate";
import { getBash } from "../vfs";
import { defineTool, toolError, toolSuccess } from "./types";

export const bashTool = defineTool({
  name: "bash",
  label: "Bash",
  description:
    "Execute bash commands in a sandboxed virtual environment. " +
    `Output is truncated to last ${DEFAULT_MAX_LINES} lines or ${DEFAULT_MAX_BYTES / 1024}KB (whichever is hit first). ` +
    "The filesystem is in-memory with user uploads in /home/user/uploads/. " +
    "Useful for: file operations (ls, cat, grep, find), text processing (awk, sed, jq, sort, uniq), " +
    "data analysis (wc, cut, paste), and general scripting. " +
    "Network access is disabled. No external runtimes (node, python, etc.) are available.",
  parameters: Type.Object({
    command: Type.String({
      description:
        "Bash command(s) to execute. Can be a single command or a script with multiple lines. " +
        "Supports pipes (|), redirections (>, >>), command chaining (&&, ||, ;), " +
        "variables, loops, conditionals, and functions.",
    }),
    explanation: Type.Optional(
      Type.String({
        description: "Brief explanation (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      const bash = getBash();
      const result = await bash.exec(params.command);

      let output = "";

      if (result.stdout) {
        output += result.stdout;
      }

      if (result.stderr) {
        if (output && !output.endsWith("\n")) output += "\n";
        output += `stderr: ${result.stderr}`;
      }

      if (result.exitCode !== 0) {
        if (output && !output.endsWith("\n")) output += "\n";
        output += `[exit code: ${result.exitCode}]`;
      }

      if (!output) {
        output = "[no output]";
      }

      output = output.trim();

      const truncation = truncateTail(output);
      let outputText = truncation.content;

      if (truncation.truncated) {
        const startLine = truncation.totalLines - truncation.outputLines + 1;
        const endLine = truncation.totalLines;
        if (truncation.truncatedBy === "lines") {
          outputText += `\n\n[Showing last ${truncation.outputLines} of ${truncation.totalLines} lines. Output truncated.]`;
        } else {
          outputText += `\n\n[Showing lines ${startLine}-${endLine} of ${truncation.totalLines} (${formatSize(DEFAULT_MAX_BYTES)} limit). Output truncated.]`;
        }
      }

      return toolSuccess({ output: outputText, exitCode: result.exitCode });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error executing bash command";
      return toolError(message);
    }
  },
});
