import { undoManager } from "../undo";
import { defineTool, toolSuccess, toolError } from "./types";

/**
 * Undo tool - Programmatically reverse Excel operations
 *
 * This provides real undo functionality by reversing tracked operations.
 * All write operations (set_cell_range, modify_object, etc.) automatically
 * register their undo actions.
 */
export const undoTool = defineTool({
  name: "undo",
  description:
    "Undo the last Excel operation(s). This programmatically reverses changes made by write operations. Use this if you made a mistake or need to revert changes.",
  parameters: {
    type: "object",
    properties: {
      steps: {
        type: "number",
        description:
          "Number of operations to undo (default: 1). Each write operation counts as one step.",
        default: 1,
      },
    },
  },
  execute: async ({ steps = 1 }) => {
    if (!undoManager.canUndo()) {
      return toolError("No operations to undo. The undo history is empty.");
    }

    const undoneOperations: string[] = [];

    try {
      for (let i = 0; i < steps; i++) {
        if (!undoManager.canUndo()) {
          break;
        }

        const description = await undoManager.undo();
        if (description) {
          undoneOperations.push(description);
        }
      }

      if (undoneOperations.length === 0) {
        return toolError("No operations were undone.");
      }

      const message = `✅ Successfully undone ${undoneOperations.length} operation${undoneOperations.length > 1 ? "s" : ""}:

${undoneOperations.map((desc, i) => `${i + 1}. ${desc}`).join("\n")}

The data has been restored to its previous state.`;

      return toolSuccess(message);
    } catch (error) {
      return toolError(
        `Failed to undo: ${error instanceof Error ? error.message : String(error)}\n\nPartially undone operations:\n${undoneOperations.map((desc, i) => `${i + 1}. ${desc}`).join("\n")}`
      );
    }
  },
});

/**
 * Get undo history tool - See what operations can be undone
 */
export const undoHistoryTool = defineTool({
  name: "undo_history",
  description:
    "View the history of operations that can be undone. Shows the most recent operations first.",
  parameters: {
    type: "object",
    properties: {},
  },
  execute: async () => {
    const history = undoManager.getUndoHistory();

    if (history.length === 0) {
      return toolSuccess("No operations in undo history.");
    }

    const historyList = history
      .reverse() // Most recent first
      .map((op, i) => {
        const date = new Date(op.timestamp);
        return `${i + 1}. ${op.description} (${date.toLocaleTimeString()})`;
      })
      .join("\n");

    return toolSuccess(
      `Undo History (${history.length} operation${history.length > 1 ? "s" : ""}):\n\n${historyList}\n\nUse the 'undo' tool to reverse these operations.`
    );
  },
});
