/**
 * Telemetry Hooks Interface
 *
 * Provides optional hooks for monitoring tool execution.
 * Implementations can integrate with any monitoring service (Datadog, Application Insights, etc.)
 */

export interface TelemetryHooks {
  /**
   * Called after a tool executes
   * @param toolCallId - Unique ID for this tool invocation
   * @param toolName - Name of the tool that was called
   * @param success - Whether the tool succeeded (parses JSON result for success field)
   * @param durationMs - Execution time in milliseconds (if available)
   * @param threwException - True if tool crashed with exception, false if returned {"success":false}
   * @param errorMessage - Error message if failed
   * @param resultPreview - First 500 chars of result for debugging
   */
  onToolResult?(params: {
    toolCallId: string;
    toolName: string;
    success: boolean;
    durationMs?: number;
    threwException?: boolean;
    errorMessage?: string;
    resultPreview?: string;
  }): void;

  /**
   * Called when user context is set (after authentication)
   */
  onUserContext?(user: { email: string; name: string; id: string }): void;

  /**
   * Called on errors
   */
  onError?(error: Error, context?: Record<string, any>): void;
}

/**
 * Global telemetry hooks instance
 * Set this during app initialization to enable monitoring
 */
export let telemetryHooks: TelemetryHooks | null = null;

/**
 * Initialize telemetry hooks
 * Call this at app startup to enable monitoring
 */
export function initTelemetryHooks(hooks: TelemetryHooks): void {
  telemetryHooks = hooks;
  console.log("[Telemetry] Hooks initialized");
}

/**
 * Helper to safely call hook
 */
function callHook<T extends keyof TelemetryHooks>(
  hookName: T,
  ...args: Parameters<NonNullable<TelemetryHooks[T]>>
): void {
  try {
    const hook = telemetryHooks?.[hookName];
    if (hook && typeof hook === "function") {
      // @ts-expect-error - TypeScript struggles with spread args on union types
      hook(...args);
    }
  } catch (error) {
    console.error(`[Telemetry] Hook ${hookName} failed:`, error);
  }
}

/**
 * Log tool execution result
 */
export function logToolCall(params: {
  toolCallId: string;
  toolName: string;
  success: boolean;
  durationMs?: number;
  threwException?: boolean;
  errorMessage?: string;
  resultPreview?: string;
}): void {
  callHook("onToolResult", params);
}

/**
 * Log user context
 */
export function logUserContext(user: { email: string; name: string; id: string }): void {
  callHook("onUserContext", user);
}

/**
 * Log error
 */
export function logError(error: Error, context?: Record<string, any>): void {
  callHook("onError", error, context);
}
