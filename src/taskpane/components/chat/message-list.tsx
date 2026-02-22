import { code } from "@streamdown/code";
import {
  Brain,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  Edit3,
  Loader2,
  Wrench,
  XCircle,
} from "lucide-react";
import type { AnchorHTMLAttributes } from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Streamdown } from "streamdown";
import { type DirtyRange, mergeRanges } from "../../../lib/dirty-tracker";
import { navigateTo } from "../../../lib/excel/api";
import type { ChatMessage, MessagePart } from "../../../lib/message-utils";
import { useChat } from "./chat-context";

function ThinkingBlock({
  thinking,
  isStreaming,
}: {
  thinking: string;
  isStreaming?: boolean;
}) {
  const [isExpanded, setIsExpanded] = useState(false);

  return (
    <div className="mb-2 border border-(--chat-border) bg-(--chat-bg) rounded-sm overflow-hidden">
      <button
        type="button"
        onClick={() => setIsExpanded(!isExpanded)}
        className="w-full flex items-center gap-1.5 px-2 py-1 text-[10px] uppercase tracking-wider text-(--chat-accent) hover:bg-(--chat-bg-secondary) transition-colors"
      >
        {isExpanded ? <ChevronDown size={10} /> : <ChevronRight size={10} />}
        <Brain size={10} />
        thinking
        {isStreaming && <span className="animate-pulse ml-1">...</span>}
      </button>
      {isExpanded && (
        <div className="px-2 py-1.5 text-xs text-(--chat-text-muted) whitespace-pre-wrap wrap-break-word border-t border-(--chat-border) max-h-20 overflow-y-auto">
          {thinking}
        </div>
      )}
    </div>
  );
}

type ToolCallPart = Extract<MessagePart, { type: "toolCall" }>;

function parseDirtyRanges(result: string | undefined): DirtyRange[] | null {
  if (!result) return null;
  try {
    const parsed = JSON.parse(result);
    if (parsed._dirtyRanges && Array.isArray(parsed._dirtyRanges)) {
      return parsed._dirtyRanges;
    }
  } catch {
    // Not valid JSON or no dirty ranges
  }
  return null;
}

function DirtyRangeLink({ range }: { range: DirtyRange }) {
  const { getSheetName } = useChat();
  const sheetName = getSheetName(range.sheetId);

  if (range.sheetId < 0) {
    const label =
      range.range === "*" ? "Unknown sheet" : `Unknown!${range.range}`;
    return <span className="text-(--chat-warning-muted)">{label}</span>;
  }

  if (!sheetName) {
    return null;
  }

  const label =
    range.range === "*" ? `${sheetName} (all)` : `${sheetName}!${range.range}`;

  return (
    <button
      type="button"
      className="text-(--chat-warning) hover:underline cursor-pointer"
      onClick={(e) => {
        e.stopPropagation();
        const navRange = range.range === "*" ? undefined : range.range;
        navigateTo(range.sheetId, navRange).catch((err) => {
          console.error("[DirtyRange] Navigation failed:", err);
        });
      }}
    >
      {label}
    </button>
  );
}

function DirtyRangeLinks({ ranges }: { ranges: DirtyRange[] }) {
  const { getSheetName } = useChat();
  const merged = useMemo(() => mergeRanges(ranges), [ranges]);

  const validRanges = merged.filter(
    (r) => r.sheetId < 0 || getSheetName(r.sheetId),
  );

  if (validRanges.length === 0) return null;

  return (
    <>
      {validRanges.map((r, i) => (
        <span key={`${r.sheetId}-${r.range}`}>
          {i > 0 && <span className="text-(--chat-warning-muted)">, </span>}
          <DirtyRangeLink range={r} />
        </span>
      ))}
    </>
  );
}

function DirtyRangeSummary({ ranges }: { ranges: DirtyRange[] }) {
  const { getSheetName, state } = useChat();
  const merged = useMemo(() => mergeRanges(ranges), [ranges]);

  console.log("[DirtyRangeSummary] sheetNames map:", state.sheetNames);
  console.log("[DirtyRangeSummary] merged ranges:", merged);

  if (merged.length === 0) return null;

  const formatBrief = (range: DirtyRange): string | null => {
    if (range.sheetId < 0) return range.range === "*" ? "unknown" : range.range;
    const sheetName = getSheetName(range.sheetId);
    console.log(
      `[DirtyRangeSummary] sheetId=${range.sheetId} -> sheetName=${sheetName}`,
    );
    if (!sheetName) return null;
    if (range.range === "*") return sheetName;
    return `${range.range}`;
  };

  if (merged.length === 1) {
    const brief = formatBrief(merged[0]);
    if (!brief) return null;
    return (
      <span className="text-[10px] text-(--chat-warning) truncate">
        → {brief}
      </span>
    );
  }

  const validRanges = merged.filter(
    (r) => r.sheetId < 0 || getSheetName(r.sheetId),
  );
  if (validRanges.length === 0) return null;

  return (
    <span className="text-[10px] text-(--chat-warning)">
      → {validRanges.length} ranges
    </span>
  );
}

function ToolCallBlock({ part }: { part: ToolCallPart }) {
  const { getSheetName } = useChat();
  const [isExpanded, setIsExpanded] = useState(false);
  const explanation = (part.args as { explanation?: string })?.explanation;

  const dirtyRanges = useMemo(
    () => parseDirtyRanges(part.result),
    [part.result],
  );
  const hasValidDirtyRanges = useMemo(() => {
    if (!dirtyRanges || dirtyRanges.length === 0) return false;
    return dirtyRanges.some((r) => r.sheetId < 0 || getSheetName(r.sheetId));
  }, [dirtyRanges, getSheetName]);

  const statusIcon = {
    pending: (
      <Loader2 size={10} className="animate-spin text-(--chat-text-muted)" />
    ),
    running: (
      <Loader2 size={10} className="animate-spin text-(--chat-accent)" />
    ),
    complete: <CheckCircle2 size={10} className="text-green-500" />,
    error: <XCircle size={10} className="text-red-500" />,
  }[part.status];

  return (
    <div className="mt-3 mb-2 border border-(--chat-border) bg-(--chat-bg) rounded-sm overflow-hidden">
      <button
        type="button"
        onClick={() => setIsExpanded(!isExpanded)}
        className={`w-full flex items-center gap-1.5 px-2 py-1 text-[10px] tracking-wider text-(--chat-text-secondary) hover:bg-(--chat-bg-secondary) transition-colors ${explanation ? "normal-case" : "uppercase"}`}
      >
        {isExpanded ? <ChevronDown size={10} /> : <ChevronRight size={10} />}
        <Wrench size={10} />
        <span className="flex-1 text-left font-medium truncate">
          {explanation || part.name}
        </span>
        {hasValidDirtyRanges && !isExpanded && (
          <span
            className="flex items-center gap-1.5 text-(--chat-warning) shrink-0"
            title="Modified cells"
          >
            <Edit3 size={9} />
            <DirtyRangeSummary ranges={dirtyRanges} />
          </span>
        )}
        {statusIcon}
      </button>
      {isExpanded && (
        <div className="border-t border-(--chat-border)">
          {hasValidDirtyRanges && (
            <div className="px-2 py-1 text-[10px] bg-(--chat-warning-bg) text-(--chat-warning) flex items-center gap-1 flex-wrap">
              <Edit3 size={9} className="shrink-0" />
              <span className="shrink-0">Modified:</span>
              <DirtyRangeLinks ranges={dirtyRanges} />
            </div>
          )}
          <div className="px-2 py-1.5 text-xs">
            <div className="text-(--chat-text-muted) text-[10px] uppercase mb-1">
              args
            </div>
            <div className="markdown-content max-h-32 overflow-y-auto **:data-[streamdown=code-block]:my-0 **:data-[streamdown=code-block]:border-0">
              <Streamdown
                plugins={{ code }}
              >{`\`\`\`json\n${JSON.stringify(part.args, null, 2)}\n\`\`\``}</Streamdown>
            </div>
          </div>
          {part.images && part.images.length > 0 && (
            <div className="px-2 py-1.5 border-t border-(--chat-border)">
              {part.images.map((img, imgIdx) => (
                <img
                  key={`${part.id}-img-${imgIdx}`}
                  src={`data:${img.mimeType};base64,${img.data}`}
                  alt={`Tool result ${imgIdx + 1}`}
                  className="max-w-full rounded-sm border border-(--chat-border)"
                />
              ))}
            </div>
          )}
          {part.result && (
            <div className="px-2 py-1.5 text-xs border-t border-(--chat-border)">
              <div className="text-(--chat-text-muted) text-[10px] uppercase mb-1">
                {part.status === "error" ? "error" : "result"}
              </div>
              <div
                className={`markdown-content max-h-40 overflow-y-auto **:data-[streamdown=code-block]:my-0 **:data-[streamdown=code-block]:border-0 ${part.status === "error" ? "[&_code]:text-red-400!" : ""}`}
              >
                <Streamdown
                  plugins={{ code }}
                >{`\`\`\`json\n${part.result}\n\`\`\``}</Streamdown>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function LoadingIndicator() {
  return (
    <div
      className="flex items-center gap-2 text-(--chat-text-muted) text-sm"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      <Loader2 size={14} className="animate-spin" />
      <span>thinking...</span>
    </div>
  );
}

function parseCitationUri(
  href: string,
): { sheetId: number; range?: string } | null {
  if (!href.startsWith("#cite:")) return null;
  const path = href.slice("#cite:".length);
  const bangIdx = path.indexOf("!");
  if (bangIdx === -1) {
    const sheetId = Number.parseInt(path, 10);
    return Number.isNaN(sheetId) ? null : { sheetId };
  }
  const sheetId = Number.parseInt(path.slice(0, bangIdx), 10);
  const range = path.slice(bangIdx + 1);
  return Number.isNaN(sheetId) ? null : { sheetId, range };
}

function CitationLink({
  href,
  children,
  ...props
}: AnchorHTMLAttributes<HTMLAnchorElement>) {
  const citation = href ? parseCitationUri(href) : null;

  if (!citation) {
    return (
      <a href={href} target="_blank" rel="noopener noreferrer" {...props}>
        {children}
      </a>
    );
  }

  return (
    <button
      type="button"
      className="text-(--chat-accent) hover:underline cursor-pointer"
      onClick={() => {
        navigateTo(citation.sheetId, citation.range).catch((err) => {
          console.error("[Citation] Navigation failed:", err);
        });
      }}
    >
      {children}
    </button>
  );
}

const markdownComponents = { a: CitationLink };

function MarkdownContent({
  text,
  isAnimating,
}: {
  text: string;
  isAnimating?: boolean;
}) {
  return (
    <div className="markdown-content">
      <Streamdown
        plugins={{ code }}
        components={markdownComponents}
        isAnimating={isAnimating}
      >
        {text}
      </Streamdown>
    </div>
  );
}

function renderParts(
  parts: MessagePart[],
  isStreaming: boolean,
  messageId: string,
) {
  const lastPart = parts[parts.length - 1];
  const isStreamingThinking = isStreaming && lastPart?.type === "thinking";
  const isStreamingText = isStreaming && lastPart?.type === "text";

  return parts.map((part, idx) => {
    const key =
      part.type === "toolCall" ? part.id : `${messageId}-${part.type}-${idx}`;
    const isLastPart = idx === parts.length - 1;
    if (part.type === "thinking") {
      return (
        <ThinkingBlock
          key={key}
          thinking={part.thinking}
          isStreaming={isStreamingThinking && isLastPart}
        />
      );
    }
    if (part.type === "toolCall") {
      return <ToolCallBlock key={key} part={part} />;
    }
    return (
      <MarkdownContent
        key={key}
        text={part.text}
        isAnimating={isStreamingText && isLastPart}
      />
    );
  });
}

function UserBubble({ message }: { message: ChatMessage }) {
  return (
    <div
      className="ml-8 px-3 py-2 text-sm leading-relaxed bg-(--chat-user-bg) border border-(--chat-border)"
      style={{
        borderRadius: "var(--chat-radius)",
        fontFamily: "var(--chat-font-mono)",
      }}
    >
      {renderParts(message.parts, false, message.id)}
    </div>
  );
}

function AssistantBubble({
  messages,
  isStreaming,
}: {
  messages: ChatMessage[];
  isStreaming: boolean;
}) {
  const allParts: { part: MessagePart; messageId: string; isLast: boolean }[] =
    [];
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    const isLastMessage = i === messages.length - 1;
    for (let j = 0; j < msg.parts.length; j++) {
      allParts.push({
        part: msg.parts[j],
        messageId: msg.id,
        isLast: isLastMessage && j === msg.parts.length - 1,
      });
    }
  }

  return (
    <div
      className="text-sm leading-relaxed"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      {allParts.map(({ part, messageId, isLast }, idx) => {
        const key =
          part.type === "toolCall"
            ? part.id
            : `${messageId}-${part.type}-${idx}`;
        if (part.type === "thinking") {
          return (
            <ThinkingBlock
              key={key}
              thinking={part.thinking}
              isStreaming={isStreaming && isLast}
            />
          );
        }
        if (part.type === "toolCall") {
          return <ToolCallBlock key={key} part={part} />;
        }
        return (
          <MarkdownContent
            key={key}
            text={part.text}
            isAnimating={isStreaming && isLast && part.type === "text"}
          />
        );
      })}
      {isStreaming && allParts.length === 0 && (
        <span className="animate-pulse">▊</span>
      )}
    </div>
  );
}

type MessageGroup =
  | { type: "user"; message: ChatMessage }
  | { type: "assistant"; messages: ChatMessage[] };

function groupMessages(messages: ChatMessage[]): MessageGroup[] {
  const groups: MessageGroup[] = [];
  let currentAssistantGroup: ChatMessage[] = [];

  for (const msg of messages) {
    if (msg.role === "user") {
      if (currentAssistantGroup.length > 0) {
        groups.push({ type: "assistant", messages: currentAssistantGroup });
        currentAssistantGroup = [];
      }
      groups.push({ type: "user", message: msg });
    } else {
      currentAssistantGroup.push(msg);
    }
  }

  if (currentAssistantGroup.length > 0) {
    groups.push({ type: "assistant", messages: currentAssistantGroup });
  }

  return groups;
}

export function MessageList() {
  const { state } = useChat();
  const containerRef = useRef<HTMLDivElement>(null);
  const shouldAutoScroll = useRef(true);

  const handleScroll = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;
    const { scrollTop, scrollHeight, clientHeight } = container;
    const distanceFromBottom = scrollHeight - scrollTop - clientHeight;
    shouldAutoScroll.current = distanceFromBottom < 100;
  }, []);

  // biome-ignore lint/correctness/useExhaustiveDependencies: intentional - trigger scroll on message/streaming changes
  useEffect(() => {
    if (containerRef.current && shouldAutoScroll.current) {
      containerRef.current.scrollTop = containerRef.current.scrollHeight;
    }
  }, [state.messages, state.isStreaming]);

  if (state.messages.length === 0) {
    return (
      <div
        className="flex-1 flex flex-col items-center justify-center p-6 text-center"
        style={{ fontFamily: "var(--chat-font-mono)" }}
      >
        <div className="text-(--chat-text-muted) text-xs uppercase tracking-widest mb-2">
          no messages
        </div>
        <div className="text-(--chat-text-secondary) text-sm max-w-[200px]">
          Start a conversation to interact with your Excel data
        </div>
      </div>
    );
  }

  const groups = groupMessages(state.messages);
  const lastMessage = state.messages[state.messages.length - 1];
  const showLoading = state.isStreaming && lastMessage?.role === "user";
  const lastGroup = groups[groups.length - 1];
  const isStreamingAssistant =
    state.isStreaming && lastGroup?.type === "assistant";

  return (
    <div
      ref={containerRef}
      onScroll={handleScroll}
      className="flex-1 overflow-y-auto p-3 space-y-3"
      style={{
        scrollbarWidth: "thin",
        scrollbarColor: "var(--chat-scrollbar) transparent",
      }}
    >
      {groups.map((group, i) => {
        if (group.type === "user") {
          return <UserBubble key={group.message.id} message={group.message} />;
        }
        const groupKey = group.messages.map((m) => m.id).join("-");
        return (
          <AssistantBubble
            key={groupKey}
            messages={group.messages}
            isStreaming={isStreamingAssistant && i === groups.length - 1}
          />
        );
      })}
      {showLoading && <LoadingIndicator />}
    </div>
  );
}
