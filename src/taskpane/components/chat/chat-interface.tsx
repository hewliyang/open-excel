import {
  Check,
  ChevronDown,
  Eye,
  EyeOff,
  MessageSquare,
  Moon,
  Plus,
  Settings,
  Sun,
  Trash2,
  Upload,
} from "lucide-react";
import { type DragEvent, type ReactNode, useCallback, useEffect, useRef, useState } from "react";
import { ChatProvider, useChat } from "./chat-context";
import { ChatInput } from "./chat-input";
import { MessageList } from "./message-list";
import { SettingsPanel } from "./settings-panel";
import type { ChatTab } from "./types";

type Theme = "light" | "dark";
const THEME_KEY = "openexcel-theme";

function useTheme() {
  const [theme, setTheme] = useState<Theme>(() => {
    const saved = localStorage.getItem(THEME_KEY) as Theme | null;
    const initial = saved ?? (window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark");
    document.documentElement.setAttribute("data-theme", initial);
    return initial;
  });

  const toggle = () => {
    const next = theme === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem(THEME_KEY, next);
    setTheme(next);
  };

  return { theme, toggle };
}

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}k`;
  return n.toString();
}

function formatCost(n: number): string {
  if (n < 0.01) return `$${n.toFixed(4)}`;
  return `$${n.toFixed(3)}`;
}

function StatsBar() {
  const { state } = useChat();
  const { sessionStats, providerConfig } = state;

  if (!providerConfig) return null;

  const totalTokens = sessionStats.inputTokens + sessionStats.outputTokens;
  const contextPct =
    sessionStats.contextWindow > 0 ? ((totalTokens / sessionStats.contextWindow) * 100).toFixed(1) : "0";

  return (
    <div
      className="flex items-center justify-between px-3 py-1.5 text-[10px] border-t border-(--chat-border) bg-(--chat-bg-secondary) text-(--chat-text-muted)"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      <div className="flex items-center gap-3">
        <span title="Input tokens">↑{formatTokens(sessionStats.inputTokens)}</span>
        <span title="Output tokens">↓{formatTokens(sessionStats.outputTokens)}</span>
        {sessionStats.cacheRead > 0 && <span title="Cache read tokens">R{formatTokens(sessionStats.cacheRead)}</span>}
        {sessionStats.cacheWrite > 0 && (
          <span title="Cache write tokens">W{formatTokens(sessionStats.cacheWrite)}</span>
        )}
        <span title="Total cost">{formatCost(sessionStats.totalCost)}</span>
        {sessionStats.contextWindow > 0 && (
          <span title="Context usage">
            {contextPct}%/{formatTokens(sessionStats.contextWindow)}
          </span>
        )}
      </div>
      <div className="flex items-center gap-1">
        <span>{providerConfig.provider}</span>
        <span className="text-(--chat-text-secondary)">{providerConfig.model}</span>
        {providerConfig.thinking !== "none" && (
          <span className="text-(--chat-accent)">• {providerConfig.thinking}</span>
        )}
      </div>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: ReactNode }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`
        flex items-center gap-1.5 px-3 py-2 text-xs uppercase tracking-wider
        border-b-2 transition-colors
        ${
          active
            ? "border-(--chat-accent) text-(--chat-text-primary)"
            : "border-transparent text-(--chat-text-muted) hover:text-(--chat-text-secondary)"
        }
      `}
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      {children}
    </button>
  );
}

function SessionDropdown({ onSelect }: { onSelect: () => void }) {
  const { state, newSession, switchSession, deleteCurrentSession } = useChat();
  const [open, setOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const isStreaming = state.isStreaming;

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    if (open) document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [open]);

  const currentName = state.currentSession?.name ?? "New Chat";
  const truncatedName = currentName.length > 20 ? `${currentName.slice(0, 18)}…` : currentName;

  const handleNewSession = async () => {
    console.log("[UI] handleNewSession clicked");
    await newSession();
    console.log("[UI] newSession completed");
    setOpen(false);
    onSelect();
  };

  const handleSwitch = async (id: string) => {
    await switchSession(id);
    setOpen(false);
    onSelect();
  };

  return (
    <div className="relative" ref={dropdownRef}>
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className={`
          flex items-center gap-1 px-3 py-2 text-xs uppercase tracking-wider
          border-b-2 border-(--chat-accent) text-(--chat-text-primary) transition-colors
        `}
        style={{ fontFamily: "var(--chat-font-mono)" }}
      >
        <MessageSquare size={12} />
        <span className="max-w-[100px] truncate">{truncatedName}</span>
        <ChevronDown size={12} className={`transition-transform ${open ? "rotate-180" : ""}`} />
      </button>

      {open && (
        <div
          className="absolute top-full left-0 mt-1 w-56 bg-(--chat-bg) border border-(--chat-border) rounded shadow-lg z-50 overflow-hidden"
          style={{ fontFamily: "var(--chat-font-mono)" }}
        >
          <button
            type="button"
            onClick={handleNewSession}
            disabled={isStreaming}
            className={`w-full flex items-center gap-2 px-3 py-2 text-xs transition-colors border-b border-(--chat-border) ${
              isStreaming
                ? "text-(--chat-text-muted) cursor-not-allowed"
                : "text-(--chat-accent) hover:bg-(--chat-bg-secondary)"
            }`}
          >
            <Plus size={14} />
            New Chat
          </button>

          <div className="max-h-48 overflow-y-auto">
            {state.sessions.map((session) => {
              const isCurrent = session.id === state.currentSession?.id;
              const isDisabled = isStreaming && !isCurrent;
              return (
                <button
                  type="button"
                  key={session.id}
                  disabled={isDisabled}
                  className={`
                    flex items-center justify-between px-3 py-2 text-xs transition-colors w-full text-left
                    ${isCurrent ? "bg-(--chat-bg-secondary)" : ""}
                    ${isDisabled ? "cursor-not-allowed opacity-50" : "cursor-pointer hover:bg-(--chat-bg-secondary)"}
                  `}
                  onClick={() => handleSwitch(session.id)}
                >
                  <div className="flex items-center gap-2 min-w-0 flex-1">
                    {session.id === state.currentSession?.id ? (
                      <Check size={12} className="text-(--chat-accent) shrink-0" />
                    ) : (
                      <div className="w-3 shrink-0" />
                    )}
                    <span className="truncate text-(--chat-text-primary)">{session.name}</span>
                  </div>
                  <span className="text-[10px] text-(--chat-text-muted) shrink-0 ml-2">{session.messages.length}</span>
                </button>
              );
            })}
          </div>

          {state.sessions.length > 1 && state.currentSession && (
            <button
              type="button"
              disabled={isStreaming}
              onClick={async (e) => {
                e.stopPropagation();
                await deleteCurrentSession();
                setOpen(false);
              }}
              className={`w-full flex items-center gap-2 px-3 py-2 text-xs transition-colors border-t border-(--chat-border) ${
                isStreaming
                  ? "text-(--chat-text-muted) cursor-not-allowed"
                  : "text-(--chat-error) hover:bg-(--chat-bg-secondary)"
              }`}
            >
              <Trash2 size={14} />
              Delete Current Session
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function ChatHeader({
  activeTab,
  onTabChange,
  theme,
  onThemeToggle,
}: {
  activeTab: ChatTab;
  onTabChange: (tab: ChatTab) => void;
  theme: Theme;
  onThemeToggle: () => void;
}) {
  const { clearMessages, state, toggleFollowMode } = useChat();
  const followMode = state.providerConfig?.followMode ?? true;

  return (
    <div className="border-b border-(--chat-border) bg-(--chat-bg)">
      <div className="flex items-center justify-between px-2">
        <div className="flex">
          {activeTab === "chat" ? (
            <SessionDropdown onSelect={() => onTabChange("chat")} />
          ) : (
            <TabButton active={false} onClick={() => onTabChange("chat")}>
              <MessageSquare size={12} />
              Chat
            </TabButton>
          )}
          <TabButton active={activeTab === "settings"} onClick={() => onTabChange("settings")}>
            <Settings size={12} />
            Settings
          </TabButton>
        </div>
        <div className="flex items-center">
          {activeTab === "chat" && (
            <button
              type="button"
              onClick={toggleFollowMode}
              className={`p-1.5 transition-colors ${
                followMode
                  ? "text-(--chat-accent) hover:text-(--chat-text-primary)"
                  : "text-(--chat-text-muted) hover:text-(--chat-text-primary)"
              }`}
              title={followMode ? "Follow mode: ON - Click to disable" : "Follow mode: OFF - Click to enable"}
            >
              {followMode ? <Eye size={14} /> : <EyeOff size={14} />}
            </button>
          )}
          <button
            type="button"
            onClick={onThemeToggle}
            className="p-1.5 text-(--chat-text-muted) hover:text-(--chat-text-primary) transition-colors"
            title={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
          >
            {theme === "dark" ? <Sun size={14} /> : <Moon size={14} />}
          </button>
          {activeTab === "chat" && state.messages.length > 0 && (
            <button
              type="button"
              onClick={clearMessages}
              className="p-1.5 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
              title="Clear messages"
            >
              <Trash2 size={14} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function ChatContent() {
  const [activeTab, setActiveTab] = useState<ChatTab>("chat");
  const { theme, toggle } = useTheme();
  const { processFiles } = useChat();
  const [isDragOver, setIsDragOver] = useState(false);
  const dragCounterRef = useRef(0);

  const handleDragEnter = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current++;
    if (e.dataTransfer.types.includes("Files")) {
      setIsDragOver(true);
    }
  }, []);

  const handleDragLeave = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current--;
    if (dragCounterRef.current === 0) {
      setIsDragOver(false);
    }
  }, []);

  const handleDragOver = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      dragCounterRef.current = 0;
      setIsDragOver(false);
      const files = Array.from(e.dataTransfer.files);
      if (files.length > 0) {
        processFiles(files);
      }
    },
    [processFiles],
  );

  return (
    <div
      role="application"
      className="flex flex-col h-full bg-(--chat-bg) relative"
      style={{ fontFamily: "var(--chat-font-mono)" }}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDragOver={handleDragOver}
      onDrop={handleDrop}
    >
      <ChatHeader activeTab={activeTab} onTabChange={setActiveTab} theme={theme} onThemeToggle={toggle} />
      {activeTab === "chat" ? (
        <>
          <MessageList />
          <ChatInput />
          <StatsBar />
        </>
      ) : (
        <SettingsPanel />
      )}

      {/* Drag-and-drop overlay */}
      {isDragOver && (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-(--chat-bg)/80 backdrop-blur-sm">
          <div className="flex flex-col items-center gap-3 p-8 border-2 border-dashed border-(--chat-accent) rounded-lg">
            <Upload size={32} className="text-(--chat-accent)" />
            <span className="text-sm text-(--chat-text-primary)">Drop files here</span>
          </div>
        </div>
      )}
    </div>
  );
}

export function ChatInterface() {
  return (
    <ChatProvider>
      <ChatContent />
    </ChatProvider>
  );
}
