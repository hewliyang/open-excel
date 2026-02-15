import type { ErrorInfo, ReactNode } from "react";
import { Component } from "react";

interface ErrorBoundaryProps {
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  errorMessage: string;
}

export class ErrorBoundary extends Component<
  ErrorBoundaryProps,
  ErrorBoundaryState
> {
  state: ErrorBoundaryState = {
    hasError: false,
    errorMessage: "",
  };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return {
      hasError: true,
      errorMessage: error.message || "Something went wrong",
    };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("[UI] Unhandled render error:", error, errorInfo);
  }

  private handleRetry = () => {
    this.setState({ hasError: false, errorMessage: "" });
  };

  private handleReload = () => {
    window.location.reload();
  };

  render() {
    if (!this.state.hasError) {
      return this.props.children;
    }

    return (
      <div
        className="h-screen w-full flex items-center justify-center bg-(--chat-bg) px-4"
        style={{ fontFamily: "var(--chat-font-mono)" }}
      >
        <div className="w-full max-w-xl border border-(--chat-border) bg-(--chat-bg-secondary) p-4 space-y-3">
          <div className="text-xs uppercase tracking-widest text-(--chat-text-muted)">
            Unhandled UI Error
          </div>
          <div className="text-sm text-(--chat-text-primary)">
            The chat UI hit an unexpected error.
          </div>
          <pre className="max-h-48 overflow-auto text-xs text-(--chat-error) bg-(--chat-bg) border border-(--chat-border) p-2 whitespace-pre-wrap break-words">
            {this.state.errorMessage}
          </pre>
          <div className="flex gap-2">
            <button
              type="button"
              onClick={this.handleRetry}
              className="px-3 py-1.5 text-xs border border-(--chat-border) text-(--chat-text-primary) hover:bg-(--chat-bg)"
            >
              Try again
            </button>
            <button
              type="button"
              onClick={this.handleReload}
              className="px-3 py-1.5 text-xs border border-(--chat-error) text-(--chat-error) hover:bg-(--chat-error) hover:text-(--chat-bg)"
            >
              Reload add-in
            </button>
          </div>
        </div>
      </div>
    );
  }
}
