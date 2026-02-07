import { Type } from "@sinclair/typebox";
import { fileExists, getFileType, listUploads, readFileBuffer, toBase64 } from "../vfs";
import { defineTool, toolError, toolText } from "./types";

export const readTool = defineTool({
  name: "read",
  label: "Read",
  description:
    "Read a file from the virtual filesystem. " +
    "Files are uploaded by the user to /home/user/uploads/. " +
    "For images (png, jpg, gif, webp), returns the image for you to analyze visually. " +
    "For text files, returns the content directly. " +
    "Use 'bash ls /home/user/uploads' to see available files.",
  parameters: Type.Object({
    path: Type.String({
      description:
        "Path to the file. Can be absolute (starting with /) or relative to /home/user/uploads/. Example: 'image.png' or '/home/user/uploads/data.csv'",
    }),
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
        // Include text note alongside image (matches pi coding agent pattern)
        return {
          content: [
            { type: "text" as const, text: `Read image file: ${filename} [${mimeType}]` },
            { type: "image" as const, data: base64, mimeType },
          ],
          details: undefined,
        };
      }

      // For text files, return content directly
      const data = await readFileBuffer(fullPath);
      const decoder = new TextDecoder();
      const text = decoder.decode(data);
      return toolText(text);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error reading file";
      return toolError(message);
    }
  },
});
