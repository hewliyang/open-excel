import { Paperclip, Send, Square, X } from "lucide-react";
import { type ChangeEvent, type KeyboardEvent, useCallback, useRef, useState } from "react";
import { useChat } from "./chat-context";

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}

export function ChatInput() {
  const { sendMessage, state, abort, processFiles, removeUpload } = useChat();
  const [input, setInput] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const uploads = state.uploads;
  const isUploading = state.isUploading;

  const handleSubmit = useCallback(async () => {
    const trimmed = input.trim();
    if (!trimmed || state.isStreaming) return;
    const attachmentNames = uploads.map((u) => u.name);
    setInput("");
    await sendMessage(trimmed, attachmentNames.length > 0 ? attachmentNames : undefined);
  }, [input, state.isStreaming, sendMessage, uploads]);

  const handleKeyDown = useCallback(
    (e: KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSubmit();
      }
    },
    [handleSubmit],
  );

  const handleFileSelect = useCallback(
    async (e: ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (!files || files.length === 0) return;
      await processFiles(Array.from(files));
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    },
    [processFiles],
  );

  const openFilePicker = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  return (
    <div className="border-t border-(--chat-border) p-3 bg-(--chat-bg)" style={{ fontFamily: "var(--chat-font-mono)" }}>
      {state.error && <div className="text-(--chat-error) text-xs mb-2 px-1">{state.error}</div>}

      {/* Uploaded files chips */}
      {uploads.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mb-2">
          {uploads.map((file) => (
            <div
              key={file.name}
              className="flex items-center gap-1 px-2 py-1 text-[10px] bg-(--chat-bg-secondary) border border-(--chat-border) text-(--chat-text-secondary)"
              style={{ borderRadius: "var(--chat-radius)" }}
            >
              <span className="max-w-[120px] truncate" title={file.name}>
                {file.name}
              </span>
              {file.size > 0 && <span className="text-(--chat-text-muted)">{formatFileSize(file.size)}</span>}
              <button
                type="button"
                onClick={() => removeUpload(file.name)}
                className="ml-0.5 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
                title="Remove from list"
              >
                <X size={10} />
              </button>
            </div>
          ))}
        </div>
      )}

      <div className="flex items-end gap-2">
        {/* Hidden file input */}
        <input
          ref={fileInputRef}
          type="file"
          multiple
          onChange={handleFileSelect}
          className="hidden"
          accept="image/*,.txt,.csv,.json,.xml,.md,.html,.css,.js,.ts,.py,.sh"
        />

        {/* Upload button */}
        <button
          type="button"
          onClick={openFilePicker}
          disabled={isUploading || state.isStreaming}
          className={`
            flex items-center justify-center
            border border-(--chat-border) bg-(--chat-bg-secondary)
            text-(--chat-text-secondary)
            hover:bg-(--chat-bg-tertiary) hover:text-(--chat-text-primary)
            hover:border-(--chat-border-active)
            disabled:opacity-30 disabled:cursor-not-allowed disabled:hover:bg-(--chat-bg-secondary)
            transition-colors
          `}
          style={{ borderRadius: "var(--chat-radius)", minHeight: "32px", width: "32px" }}
          title="Upload files"
        >
          <Paperclip size={14} className={isUploading ? "animate-pulse" : ""} />
        </button>

        <textarea
          ref={textareaRef}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={state.providerConfig ? "Type a message..." : "Configure API key in settings"}
          disabled={!state.providerConfig}
          rows={1}
          className={`
            flex-1 resize-none bg-(--chat-input-bg) text-(--chat-text-primary)
            text-sm px-3 py-2 border border-(--chat-border)
            placeholder:text-(--chat-text-muted)
            focus:outline-none focus:border-(--chat-border-active)
            disabled:opacity-50 disabled:cursor-not-allowed
          `}
          style={{
            borderRadius: "var(--chat-radius)",
            fontFamily: "var(--chat-font-mono)",
            minHeight: "32px",
          }}
        />
        {state.isStreaming ? (
          <button
            type="button"
            onClick={abort}
            className={`
              flex items-center justify-center
              border border-(--chat-error) bg-(--chat-bg-secondary)
              text-(--chat-error)
              hover:bg-(--chat-error) hover:text-(--chat-bg)
              transition-colors
            `}
            style={{ borderRadius: "var(--chat-radius)", minHeight: "32px", width: "32px" }}
          >
            <Square size={14} />
          </button>
        ) : (
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!state.providerConfig || !input.trim()}
            className={`
              flex items-center justify-center
              border border-(--chat-border) bg-(--chat-bg-secondary)
              text-(--chat-text-secondary)
              hover:bg-(--chat-bg-tertiary) hover:text-(--chat-text-primary)
              hover:border-(--chat-border-active)
              disabled:opacity-30 disabled:cursor-not-allowed disabled:hover:bg-(--chat-bg-secondary)
              transition-colors
            `}
            style={{ borderRadius: "var(--chat-radius)", minHeight: "32px", width: "32px" }}
          >
            <Send size={14} />
          </button>
        )}
      </div>
    </div>
  );
}
