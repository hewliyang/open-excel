import { Type } from "@sinclair/typebox";
import { DEFAULT_MAX_BYTES, DEFAULT_MAX_LINES, formatSize, truncateHead } from "../truncate";
import { fileExists, getFileType, listUploads, readFileBuffer, toBase64 } from "../vfs";
import { defineTool, toolError, toolText } from "./types";

export const readTool = defineTool({
  name: "read",
  label: "Read",
  description:
    "Read a file from the virtual filesystem. " +
    "Files are uploaded by the user to /home/user/uploads/. " +
    "For images (png, jpg, gif, webp), returns the image for you to analyze visually. " +
    `For text files, output is truncated to ${DEFAULT_MAX_LINES} lines or ${DEFAULT_MAX_BYTES / 1024}KB (whichever is hit first). ` +
    "Use offset/limit for large files. When you need the full file, continue with offset until complete. " +
    "Use 'bash ls /home/user/uploads' to see available files.",
  parameters: Type.Object({
    path: Type.String({
      description:
        "Path to the file. Can be absolute (starting with /) or relative to /home/user/uploads/. Example: 'image.png' or '/home/user/uploads/data.csv'",
    }),
    offset: Type.Optional(
      Type.Number({
        description: "Line number to start reading from (1-indexed)",
      }),
    ),
    limit: Type.Optional(
      Type.Number({
        description: "Maximum number of lines to read",
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      const path = params.path;
      const fullPath = path.startsWith("/") ? path : `/home/user/uploads/${path}`;

      // Check if file exists
      if (!(await fileExists(fullPath))) {
        // List available files to help the user
        const uploads = await listUploads();
        const hint = uploads.length > 0 ? `Available files: ${uploads.join(", ")}` : "No files uploaded yet.";
        return toolError(`File not found: ${fullPath}. ${hint}`);
      }

      const filename = fullPath.split("/").pop() || "";
      const { isImage, mimeType } = getFileType(filename);

      if (isImage) {
        // Return image as base64 for vision models
        const data = await readFileBuffer(fullPath);
        const base64 = toBase64(data);
        return {
          content: [
            { type: "text" as const, text: `Read image file: ${filename} [${mimeType}]` },
            { type: "image" as const, data: base64, mimeType },
          ],
          details: undefined,
        };
      }

      // For text files, apply truncation
      const data = await readFileBuffer(fullPath);
      const decoder = new TextDecoder();
      const text = decoder.decode(data);

      const allLines = text.split("\n");
      const totalFileLines = allLines.length;

      // Apply offset (1-indexed to 0-indexed)
      const startLine = params.offset ? Math.max(0, params.offset - 1) : 0;
      const startLineDisplay = startLine + 1;

      if (startLine >= allLines.length) {
        return toolError(`Offset ${params.offset} is beyond end of file (${allLines.length} lines total)`);
      }

      // Apply limit param if specified
      let selectedContent: string;
      let userLimitedLines: number | undefined;

      if (params.limit !== undefined) {
        const endLine = Math.min(startLine + params.limit, allLines.length);
        selectedContent = allLines.slice(startLine, endLine).join("\n");
        userLimitedLines = endLine - startLine;
      } else {
        selectedContent = allLines.slice(startLine).join("\n");
      }

      // Apply truncation
      const truncation = truncateHead(selectedContent);
      let outputText: string;

      if (truncation.truncated) {
        const endLineDisplay = startLineDisplay + truncation.outputLines - 1;
        const nextOffset = endLineDisplay + 1;
        outputText = truncation.content;

        if (truncation.truncatedBy === "lines") {
          outputText += `\n\n[Showing lines ${startLineDisplay}-${endLineDisplay} of ${totalFileLines}. Use offset=${nextOffset} to continue.]`;
        } else {
          outputText += `\n\n[Showing lines ${startLineDisplay}-${endLineDisplay} of ${totalFileLines} (${formatSize(DEFAULT_MAX_BYTES)} limit). Use offset=${nextOffset} to continue.]`;
        }
      } else if (userLimitedLines !== undefined && startLine + userLimitedLines < allLines.length) {
        const remaining = allLines.length - (startLine + userLimitedLines);
        const nextOffset = startLine + userLimitedLines + 1;
        outputText = truncation.content;
        outputText += `\n\n[${remaining} more lines in file. Use offset=${nextOffset} to continue.]`;
      } else {
        outputText = truncation.content;
      }

      return toolText(outputText);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error reading file";
      return toolError(message);
    }
  },
});
