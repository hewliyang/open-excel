/**
 * Undo Manager for Excel Operations
 *
 * Tracks all write operations and allows reversing them programmatically.
 * This provides real Ctrl+Z functionality for the AI agent.
 */

export interface UndoOperation {
  id: string;
  timestamp: number;
  description: string;
  revert: () => Promise<void>;
}

class UndoManager {
  private undoStack: UndoOperation[] = [];
  private redoStack: UndoOperation[] = [];
  private maxStackSize = 50;

  /**
   * Register an operation that can be undone
   */
  registerOperation(description: string, revert: () => Promise<void>): string {
    const operation: UndoOperation = {
      id: crypto.randomUUID(),
      timestamp: Date.now(),
      description,
      revert,
    };

    this.undoStack.push(operation);

    // Clear redo stack when new operation is added
    this.redoStack = [];

    // Limit stack size
    if (this.undoStack.length > this.maxStackSize) {
      this.undoStack.shift();
    }

    console.log(`[UndoManager] Registered: ${description}`);
    return operation.id;
  }

  /**
   * Undo the last operation
   */
  async undo(): Promise<string | null> {
    const operation = this.undoStack.pop();
    if (!operation) {
      return null;
    }

    console.log(`[UndoManager] Undoing: ${operation.description}`);

    try {
      await operation.revert();
      this.redoStack.push(operation);
      return operation.description;
    } catch (error) {
      // If revert fails, put it back on the stack
      this.undoStack.push(operation);
      throw error;
    }
  }

  /**
   * Redo the last undone operation
   */
  async redo(): Promise<string | null> {
    const operation = this.redoStack.pop();
    if (!operation) {
      return null;
    }

    console.log(`[UndoManager] Redoing: ${operation.description}`);

    // For redo, we need to re-execute the original operation
    // This is more complex and requires storing the forward action too
    // For now, we'll just move it back to undo stack
    this.undoStack.push(operation);
    return operation.description;
  }

  /**
   * Get undo history
   */
  getUndoHistory(): Array<{ description: string; timestamp: number }> {
    return this.undoStack.map((op) => ({
      description: op.description,
      timestamp: op.timestamp,
    }));
  }

  /**
   * Check if undo is available
   */
  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  /**
   * Check if redo is available
   */
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Clear all history
   */
  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}

// Global undo manager instance
export const undoManager = new UndoManager();
