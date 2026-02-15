import { Paperclip, Send, Square, X } from "lucide-react";
import {
  type ChangeEvent,
  type KeyboardEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import { useChat } from "./chat-context";

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}

const LINE_HEIGHT = 20;
const MIN_ROWS = 1;
const MAX_ROWS = 2;

export function ChatInput() {
  const { sendMessage, state, abort, processFiles, removeUpload } = useChat();
  const [input, setInput] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const uploads = state.uploads;
  const isUploading = state.isUploading;

  const autoResize = useCallback(() => {
    const ta = textareaRef.current;
    if (!ta) return;
    ta.style.height = "auto";
    const min = LINE_HEIGHT * MIN_ROWS;
    const max = LINE_HEIGHT * MAX_ROWS;
    const clamped = Math.max(min, Math.min(ta.scrollHeight, max));
    ta.style.height = `${clamped}px`;
    ta.style.overflowY = ta.scrollHeight > max ? "auto" : "hidden";
  }, []);

  useEffect(() => {
    if (!input) {
      autoResize();
    }
  }, [input, autoResize]);

  const handleInputChange = useCallback(
    (e: ChangeEvent<HTMLTextAreaElement>) => {
      setInput(e.target.value);
      autoResize();
    },
    [autoResize],
  );

  const handleSubmit = useCallback(async () => {
    const trimmed = input.trim();
    if (!trimmed || state.isStreaming) return;
    const attachmentNames = uploads.map((u) => u.name);
    setInput("");
    await sendMessage(
      trimmed,
      attachmentNames.length > 0 ? attachmentNames : undefined,
    );
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
    <div
      className="border-t border-(--chat-border) px-3 py-2 bg-(--chat-bg)"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      {state.error && (
        <div className="text-(--chat-error) text-xs mb-2 px-1">
          {state.error}
        </div>
      )}

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
              {file.size > 0 && (
                <span className="text-(--chat-text-muted)">
                  {formatFileSize(file.size)}
                </span>
              )}
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

      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        onChange={handleFileSelect}
        className="hidden"
        accept="image/*,.txt,.csv,.json,.xml,.md,.html,.css,.js,.ts,.py,.sh"
      />

      {/* Input container — border on wrapper, textarea + action row inside */}
      <div
        className="bg-(--chat-input-bg) border border-(--chat-border) focus-within:border-(--chat-border-active) transition-colors"
        style={{ borderRadius: "var(--chat-radius)" }}
      >
        <textarea
          ref={textareaRef}
          value={input}
          onChange={handleInputChange}
          onKeyDown={handleKeyDown}
          placeholder={
            state.providerConfig
              ? "Type a message..."
              : "Configure API key in settings"
          }
          disabled={!state.providerConfig}
          className={`
            w-full resize-none bg-transparent text-(--chat-text-primary)
            text-sm px-3 pt-2 pb-0 border-none outline-none
            placeholder:text-(--chat-text-muted)
            disabled:opacity-50 disabled:cursor-not-allowed
          `}
          style={{
            fontFamily: "var(--chat-font-mono)",
            lineHeight: `${LINE_HEIGHT}px`,
            height: `${LINE_HEIGHT * MIN_ROWS}px`,
          }}
        />

        {/* Action row inside the border */}
        <div className="flex items-center justify-between px-1.5 py-1">
          <button
            type="button"
            onClick={openFilePicker}
            disabled={isUploading || state.isStreaming}
            className="flex items-center justify-center w-6 h-5
                       text-(--chat-text-muted) hover:text-(--chat-text-primary)
                       disabled:opacity-30 disabled:cursor-not-allowed
                       transition-colors"
            title="Upload files"
          >
            <Paperclip
              size={13}
              className={isUploading ? "animate-pulse" : ""}
            />
          </button>

          {state.isStreaming ? (
            <button
              type="button"
              onClick={abort}
              className="flex items-center justify-center w-6 h-5
                         text-(--chat-error) hover:text-(--chat-bg) hover:bg-(--chat-error)
                         transition-colors"
              style={{ borderRadius: "var(--chat-radius)" }}
            >
              <Square size={13} />
            </button>
          ) : (
            <button
              type="button"
              onClick={handleSubmit}
              disabled={!state.providerConfig || !input.trim()}
              className="flex items-center justify-center w-6 h-5
                         text-(--chat-text-muted) hover:text-(--chat-text-primary)
                         disabled:opacity-30 disabled:cursor-not-allowed
                         transition-colors"
            >
              <Send size={13} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
