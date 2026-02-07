import { Type } from "@sinclair/typebox";
import { getBash } from "../vfs";
import { defineTool, toolError, toolSuccess } from "./types";

export const bashTool = defineTool({
  name: "bash",
  label: "Bash",
  description:
    "Execute bash commands in a sandboxed virtual environment. " +
    "The filesystem is in-memory with user uploads in /home/user/uploads/. " +
    "Useful for: file operations (ls, cat, grep, find), text processing (awk, sed, jq, sort, uniq), " +
    "data analysis (wc, cut, paste), and general scripting. " +
    "Network access is disabled. Binary execution is not supported.",
  parameters: Type.Object({
    command: Type.String({
      description:
        "Bash command(s) to execute. Can be a single command or a script with multiple lines. " +
        "Supports pipes (|), redirections (>, >>), command chaining (&&, ||, ;), " +
        "variables, loops, conditionals, and functions.",
    }),
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

      return toolSuccess({ output: output.trim(), exitCode: result.exitCode });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error executing bash command";
      return toolError(message);
    }
  },
});
