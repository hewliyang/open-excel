import type { AgentMessage } from "@mariozechner/pi-agent-core";
import type {
  AssistantMessage,
  TextContent,
  ToolCall,
  ToolResultMessage,
  UserMessage,
} from "@mariozechner/pi-ai";
import type { ChatSession } from "./storage";

interface ExportSessionHeader {
  type: "session";
  version: 1;
  format: "openexcel-session-v1";
  id: string;
  workbookId: string;
  name: string;
  createdAt: string;
  updatedAt: string;
  exportedAt: string;
}

interface ExportMessageEntry {
  type: "message";
  id: string;
  parentId: string | null;
  timestamp: string;
  message: AgentMessage;
}

interface HtmlToolCall {
  id: string;
  name: string;
  args: Record<string, unknown>;
  resultText: string;
  resultImages: { mimeType: string; data: string }[];
  isError: boolean;
}

interface HtmlMessage {
  id: string;
  role: "user" | "assistant";
  timestamp: number;
  textBlocks: string[];
  thinkingBlocks: string[];
  toolCalls: HtmlToolCall[];
}

interface HtmlPayload {
  session: {
    id: string;
    workbookId: string;
    name: string;
    createdAt: number;
    updatedAt: number;
  };
  messages: HtmlMessage[];
}

function shortId(): string {
  return crypto.randomUUID().slice(0, 8);
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toText(content: string | { type: string; text?: string }[]): string {
  if (typeof content === "string") return content;
  return content
    .filter((c) => c.type === "text")
    .map((c) => c.text ?? "")
    .join("\n");
}

function extractToolResults(
  messages: AgentMessage[],
): Map<string, ToolResultMessage> {
  const map = new Map<string, ToolResultMessage>();
  for (const msg of messages) {
    if (msg.role === "toolResult") {
      const result = msg as ToolResultMessage;
      map.set(result.toolCallId, result);
    }
  }
  return map;
}

function buildHtmlPayload(
  session: ChatSession,
  messages: AgentMessage[],
): HtmlPayload {
  const toolResults = extractToolResults(messages);
  const filtered: HtmlMessage[] = [];

  messages.forEach((message, i) => {
    const id = `m-${i + 1}`;

    if (message.role === "user") {
      const user = message as UserMessage;
      filtered.push({
        id,
        role: "user",
        timestamp: user.timestamp,
        textBlocks: [toText(user.content)],
        thinkingBlocks: [],
        toolCalls: [],
      });
      return;
    }

    if (message.role === "assistant") {
      const assistant = message as AssistantMessage;
      const textBlocks: string[] = [];
      const thinkingBlocks: string[] = [];
      const toolCalls: HtmlToolCall[] = [];

      for (const block of assistant.content) {
        if (block.type === "text") {
          textBlocks.push(block.text);
        } else if (block.type === "thinking") {
          thinkingBlocks.push(block.thinking);
        } else if (block.type === "toolCall") {
          const result = toolResults.get(block.id);
          const resultText = result
            ? result.content
                .filter((c): c is TextContent => c.type === "text")
                .map((c) => c.text)
                .join("\n")
            : "";
          const resultImages = result
            ? result.content
                .filter((c) => c.type === "image")
                .map((c) => ({ mimeType: c.mimeType, data: c.data }))
            : [];

          toolCalls.push({
            id: block.id,
            name: block.name,
            args: (block as ToolCall).arguments ?? {},
            resultText,
            resultImages,
            isError: result?.isError ?? false,
          });
        }
      }

      filtered.push({
        id,
        role: "assistant",
        timestamp: assistant.timestamp,
        textBlocks,
        thinkingBlocks,
        toolCalls,
      });
    }
  });

  return {
    session: {
      id: session.id,
      workbookId: session.workbookId,
      name: session.name,
      createdAt: session.createdAt,
      updatedAt: session.updatedAt,
    },
    messages: filtered,
  };
}

function utf8ToBase64(input: string): string {
  const bytes = new TextEncoder().encode(input);
  let binary = "";
  for (let i = 0; i < bytes.length; i++)
    binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

export function serializeSessionToJsonl(
  session: ChatSession,
  messages: AgentMessage[],
): string {
  const header: ExportSessionHeader = {
    type: "session",
    version: 1,
    format: "openexcel-session-v1",
    id: session.id,
    workbookId: session.workbookId,
    name: session.name,
    createdAt: new Date(session.createdAt).toISOString(),
    updatedAt: new Date(session.updatedAt).toISOString(),
    exportedAt: new Date().toISOString(),
  };

  const lines: string[] = [JSON.stringify(header)];
  let parentId: string | null = null;

  for (const message of messages) {
    const entry: ExportMessageEntry = {
      type: "message",
      id: shortId(),
      parentId,
      timestamp: new Date(message.timestamp).toISOString(),
      message,
    };
    parentId = entry.id;
    lines.push(JSON.stringify(entry));
  }

  return lines.join("\n");
}

export function serializeSessionToHtml(
  session: ChatSession,
  messages: AgentMessage[],
): string {
  const payload = buildHtmlPayload(session, messages);
  const payloadB64 = utf8ToBase64(JSON.stringify(payload));

  return String.raw`<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(session.name)} · OpenExcel</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github-dark.min.css" />
    <style>
      :root { --bg: #0a0a0a; --bg2: #111; --panel: #0f0f0f; --border: #2a2a2a; --text: #e8e8e8; --muted: #9a9a9a; --accent: #6366f1; --ok: #2f855a; --warn: #dc2626; }
      * { box-sizing: border-box; }
      body { margin: 0; background: var(--bg); color: var(--text); font: 13px/1.5 ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }
      .layout { display: grid; grid-template-columns: 320px 1fr; min-height: 100vh; }
      .sidebar { border-right: 1px solid var(--border); background: var(--panel); position: sticky; top: 0; height: 100vh; overflow: auto; }
      .side-head { padding: 14px; border-bottom: 1px solid var(--border); position: sticky; top: 0; background: var(--panel); z-index: 2; }
      .title { font-size: 13px; font-weight: 700; margin: 0 0 4px; }
      .meta { color: var(--muted); font-size: 11px; }
      .jump-list { padding: 8px; }
      .jump { display: block; width: 100%; text-align: left; border: 1px solid var(--border); background: var(--bg2); color: var(--text); margin-bottom: 6px; padding: 8px; text-decoration: none; border-radius: 3px; }
      .jump:hover { border-color: var(--accent); background: #141414; }
      .jump .role { font-size: 10px; color: var(--muted); text-transform: uppercase; }
      .jump .preview { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 12px; }
      .content { padding: 18px; max-width: 1000px; width: 100%; margin: 0 auto; }
      .msg { border: 1px solid var(--border); background: #101010; padding: 12px; margin-bottom: 10px; scroll-margin-top: 16px; }
      .msg.user { border-left: 3px solid var(--accent); }
      .msg.assistant { border-left: 3px solid var(--ok); }
      .msg-head { display: flex; justify-content: space-between; gap: 8px; color: #b0b0b0; font-size: 12px; margin-bottom: 8px; }
      .markdown p { margin: 0 0 8px; }
      .markdown h1, .markdown h2, .markdown h3 { margin: 12px 0 8px; font-size: 14px; }
      .markdown ul, .markdown ol { margin: 0 0 8px 18px; }
      .markdown code { background: #191919; padding: 0 4px; }
      .markdown pre { background: #0d0d0d; border: 1px solid var(--border); padding: 10px; overflow-x: auto; margin: 0 0 8px; }
      .markdown pre code { background: transparent; padding: 0; }
      .thinking { margin-top: 8px; border: 1px dashed #444; }
      .thinking > summary { cursor: pointer; padding: 6px 8px; color: #9bb3ff; }
      .thinking pre { margin: 0; padding: 8px; }
      .tool-call { border: 1px solid #333; margin-top: 8px; }
      .tool-call.error { border-color: var(--warn); }
      .tool-head { padding: 6px 8px; background: #151515; border-bottom: 1px solid var(--border); }
      .tool-args, .tool-result { margin: 0; padding: 8px; background: #0d0d0d; overflow-x: auto; }
      .tool-args code, .tool-result code { background: transparent; padding: 0; display: block; white-space: pre; }
      .tool-args, .tool-result, .markdown pre { font-size: 12px; line-height: 1.45; }
      pre code.hljs { padding: 0; background: transparent !important; }
      .hljs { background: transparent !important; color: #e8e8e8; }
      .hljs-attr, .hljs-property { color: #9cdcfe; }
      .hljs-string { color: #ce9178; }
      .hljs-number, .hljs-literal { color: #b5cea8; }
      .tool-image { display: block; max-width: 100%; margin: 8px; border: 1px solid var(--border); }
      @media (max-width: 900px) { .layout { grid-template-columns: 1fr; } .sidebar { position: static; height: auto; } }
    </style>
  </head>
  <body>
    <div class="layout">
      <aside class="sidebar">
        <div class="side-head">
          <h1 class="title">${escapeHtml(session.name.replace(/^#+\s*/, ""))}</h1>
          <div class="meta">Session ${escapeHtml(session.id)}</div>
          <div class="meta">Workbook ${escapeHtml(session.workbookId)}</div>
        </div>
        <div id="jump-list" class="jump-list"></div>
      </aside>
      <main class="content">
        <div id="messages"></div>
      </main>
    </div>

    <script id="openexcel-session-data" type="application/octet-stream">${payloadB64}</script>
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/highlight.min.js"></script>
    <script>
      (function () {
        function esc(v) {
          return String(v)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/\"/g, "&quot;")
            .replace(/'/g, "&#39;");
        }

        function decodePayload() {
          var b64 = document.getElementById("openexcel-session-data").textContent || "";
          var bin = atob(b64);
          var bytes = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          var json = new TextDecoder("utf-8").decode(bytes);
          return JSON.parse(json);
        }

        function inlineMd(text) {
          return text
            .replace(/\x60([^\x60]+)\x60/g, "<code>$1</code>")
            .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
            .replace(/\*([^*]+)\*/g, "<em>$1</em>")
            .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>');
        }

        function markdownToHtmlFallback(raw) {
          var text = esc(raw || "");

          text = text.replace(/\x60\x60\x60([a-zA-Z0-9_-]*)\n([\s\S]*?)\x60\x60\x60/g, function (_, lang, code) {
            var cls = lang ? ' class="language-' + lang + '"' : "";
            return "<pre><code" + cls + ">" + code + "</code></pre>";
          });

          var lines = text.split("\n");
          var out = [];
          var inUl = false;
          var inOl = false;

          function closeLists() {
            if (inUl) { out.push("</ul>"); inUl = false; }
            if (inOl) { out.push("</ol>"); inOl = false; }
          }

          for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (!line.trim()) {
              closeLists();
              continue;
            }

            if (/^###\s+/.test(line)) { closeLists(); out.push("<h3>" + inlineMd(line.replace(/^###\s+/, "")) + "</h3>"); continue; }
            if (/^##\s+/.test(line)) { closeLists(); out.push("<h2>" + inlineMd(line.replace(/^##\s+/, "")) + "</h2>"); continue; }
            if (/^#\s+/.test(line)) { closeLists(); out.push("<h1>" + inlineMd(line.replace(/^#\s+/, "")) + "</h1>"); continue; }

            if (/^\d+\.\s+/.test(line)) {
              if (!inOl) { closeLists(); out.push("<ol>"); inOl = true; }
              out.push("<li>" + inlineMd(line.replace(/^\d+\.\s+/, "")) + "</li>");
              continue;
            }

            if (/^[-*]\s+/.test(line)) {
              if (!inUl) { closeLists(); out.push("<ul>"); inUl = true; }
              out.push("<li>" + inlineMd(line.replace(/^[-*]\s+/, "")) + "</li>");
              continue;
            }

            closeLists();
            out.push("<p>" + inlineMd(line) + "</p>");
          }

          closeLists();
          return out.join("\n");
        }

        function normalizeMarkdown(raw) {
          return String(raw || "")
            .replace(/^[\u2013\u2014\u2022]\s+/gm, "- ")
            .replace(/\r\n/g, "\n");
        }

        function renderMarkdown(raw) {
          var source = normalizeMarkdown(raw);
          try {
            if (window.marked && typeof window.marked.parse === "function") {
              return window.marked.parse(source, { gfm: true, breaks: true });
            }
          } catch (_) {
            // fall through to fallback parser
          }
          return markdownToHtmlFallback(source);
        }

        function maybeFormatJson(text) {
          if (typeof text !== "string") return null;
          var trimmed = text.trim();
          if (!trimmed) return null;
          if (!(trimmed.startsWith("{") || trimmed.startsWith("["))) return null;
          try {
            return JSON.stringify(JSON.parse(trimmed), null, 2);
          } catch (_) {
            return null;
          }
        }

        function escapeAttr(v) {
          return String(v)
            .replace(/&/g, "&amp;")
            .replace(/\"/g, "&quot;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
        }

        function renderCodeBlock(text, language, className) {
          return '<pre class="' + className + '"><code class="language-' + escapeAttr(language || "plaintext") + '">' + esc(text) + '</code></pre>';
        }

        function fmtTs(ts) {
          try { return new Date(ts).toLocaleString(); } catch { return ""; }
        }

        function previewText(msg) {
          var text = (msg.textBlocks || []).join(" ").trim();
          if (!text && msg.toolCalls && msg.toolCalls.length > 0) {
            var names = msg.toolCalls.map(function (t) { return t.name; });
            var unique = Array.from(new Set(names));
            if (unique.length === 1) {
              return "⚙ " + unique[0] + " ×" + names.length;
            }
            return "⚙ " + unique.join(", ");
          }
          return text || "(empty)";
        }

        function renderSidebar(messages) {
          var list = document.getElementById("jump-list");
          list.innerHTML = messages
            .map(function (msg, idx) {
              return (
                '<a class="jump" href="#' + msg.id + '">' +
                  '<div class="role">' + (idx + 1) + " · " + msg.role + "</div>" +
                  '<div class="preview">' + esc(previewText(msg)).slice(0, 200) + "</div>" +
                "</a>"
              );
            })
            .join("\n");
        }

        function renderMessages(messages) {
          var container = document.getElementById("messages");
          container.innerHTML = messages
            .map(function (msg, idx) {
              var textHtml = (msg.textBlocks || [])
                .map(function (block) {
                  return '<div class="markdown">' + renderMarkdown(block) + "</div>";
                })
                .join("\n");

              var thinkingHtml = (msg.thinkingBlocks || [])
                .map(function (block) {
                  return '<details class="thinking"><summary>Thinking</summary><pre>' + esc(block) + "</pre></details>";
                })
                .join("\n");

              var toolsHtml = (msg.toolCalls || [])
                .map(function (tool) {
                  var images = (tool.resultImages || [])
                    .map(function (img) {
                      return '<img class="tool-image" src="data:' + img.mimeType + ';base64,' + img.data + '" />';
                    })
                    .join("");

                  var argsText = JSON.stringify(tool.args || {}, null, 2);
                  var resultJson = maybeFormatJson(tool.resultText || "");
                  var resultLanguage = resultJson ? "json" : "plaintext";
                  var resultText = resultJson || (tool.resultText || "");

                  return (
                    '<div class="tool-call ' + (tool.isError ? "error" : "") + '">' +
                      '<div class="tool-head">' + esc(tool.name) + "</div>" +
                      renderCodeBlock(argsText, "json", "tool-args") +
                      (tool.resultText || images
                        ? '<details><summary>Result' + (tool.isError ? " (error)" : "") + '</summary>' +
                            (resultText ? renderCodeBlock(resultText, resultLanguage, "tool-result") : "") +
                            images +
                          "</details>"
                        : "") +
                    "</div>"
                  );
                })
                .join("\n");

              return (
                '<article id="' + msg.id + '" class="msg ' + msg.role + '">' +
                  '<div class="msg-head"><div>#' + (idx + 1) + " · " + msg.role + '</div><div>' + esc(fmtTs(msg.timestamp)) + "</div></div>" +
                  textHtml +
                  thinkingHtml +
                  toolsHtml +
                "</article>"
              );
            })
            .join("\n");
        }

        function applySyntaxHighlighting() {
          if (!(window.hljs && typeof window.hljs.highlightAll === "function")) return;
          try {
            window.hljs.highlightAll();
          } catch (_) {
            // ignore highlighting failures
          }
        }

        var payload = decodePayload();
        renderSidebar(payload.messages || []);
        renderMessages(payload.messages || []);
        applySyntaxHighlighting();
      })();
    </script>
  </body>
</html>`;
}

export function downloadTextFile(
  filename: string,
  content: string,
  mimeType: string,
): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
