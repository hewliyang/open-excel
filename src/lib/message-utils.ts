import type { AgentMessage } from "@mariozechner/pi-agent-core";
import type {
  AssistantMessage,
  ImageContent,
  TextContent,
  ToolResultMessage,
  UserMessage,
} from "@mariozechner/pi-ai";

export type ToolCallStatus = "pending" | "running" | "complete" | "error";

export type MessagePart =
  | { type: "text"; text: string }
  | { type: "thinking"; thinking: string }
  | {
      type: "toolCall";
      id: string;
      name: string;
      args: Record<string, unknown>;
      status: ToolCallStatus;
      result?: string;
      images?: { data: string; mimeType: string }[];
    };

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  parts: MessagePart[];
  timestamp: number;
}

export interface SessionStats {
  inputTokens: number;
  outputTokens: number;
  cacheRead: number;
  cacheWrite: number;
  totalCost: number;
  contextWindow: number;
  lastInputTokens: number;
}

export function stripEnrichment(
  content: string | { type: string; text?: string }[],
): string {
  let text: string;
  if (typeof content === "string") {
    text = content;
  } else {
    text = content
      .filter((c) => c.type === "text")
      .map((c) => c.text ?? "")
      .join("\n");
  }
  text = text.replace(/^<attachments>\n[\s\S]*?\n<\/attachments>\n\n/, "");
  text = text.replace(/^<wb_context>\n[\s\S]*?\n<\/wb_context>\n\n/, "");
  return text;
}

export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

export function extractPartsFromAssistantMessage(
  message: AgentMessage,
  existingParts: MessagePart[] = [],
): MessagePart[] {
  if (message.role !== "assistant") return existingParts;

  const assistantMsg = message as AssistantMessage;
  const existingToolCalls = new Map<string, MessagePart>();
  for (const part of existingParts) {
    if (part.type === "toolCall") {
      existingToolCalls.set(part.id, part);
    }
  }

  return assistantMsg.content.map((block): MessagePart => {
    if (block.type === "text") {
      return { type: "text", text: block.text };
    }
    if (block.type === "thinking") {
      return { type: "thinking", thinking: block.thinking };
    }
    const existing = existingToolCalls.get(block.id);
    return {
      type: "toolCall",
      id: block.id,
      name: block.name,
      args: block.arguments as Record<string, unknown>,
      status: existing?.type === "toolCall" ? existing.status : "pending",
      result: existing?.type === "toolCall" ? existing.result : undefined,
    };
  });
}

export function agentMessagesToChatMessages(
  agentMessages: AgentMessage[],
): ChatMessage[] {
  const result: ChatMessage[] = [];
  for (const msg of agentMessages) {
    if (msg.role === "user") {
      const text = stripEnrichment((msg as UserMessage).content);
      result.push({
        id: generateId(),
        role: "user",
        parts: [{ type: "text", text }],
        timestamp: msg.timestamp,
      });
    } else if (msg.role === "assistant") {
      const parts = extractPartsFromAssistantMessage(msg);
      result.push({
        id: generateId(),
        role: "assistant",
        parts,
        timestamp: msg.timestamp,
      });
    } else if (msg.role === "toolResult") {
      const toolResult = msg as ToolResultMessage;
      for (let i = result.length - 1; i >= 0; i--) {
        const chatMsg = result[i];
        if (chatMsg.role !== "assistant") continue;
        const partIdx = chatMsg.parts.findIndex(
          (p) => p.type === "toolCall" && p.id === toolResult.toolCallId,
        );
        if (partIdx !== -1) {
          const part = chatMsg.parts[partIdx];
          if (part.type === "toolCall") {
            const resultText = toolResult.content
              .filter((c): c is TextContent => c.type === "text")
              .map((c) => c.text)
              .join("\n");
            const images = toolResult.content
              .filter((c): c is ImageContent => c.type === "image")
              .map((c) => ({ data: c.data, mimeType: c.mimeType }));
            chatMsg.parts[partIdx] = {
              ...part,
              status: toolResult.isError ? "error" : "complete",
              result: resultText,
              images: images.length > 0 ? images : undefined,
            };
          }
          break;
        }
      }
    }
  }
  return result;
}

export function deriveStats(
  agentMessages: AgentMessage[],
): Omit<SessionStats, "contextWindow"> {
  let inputTokens = 0;
  let outputTokens = 0;
  let cacheRead = 0;
  let cacheWrite = 0;
  let totalCost = 0;
  let lastInputTokens = 0;
  for (const msg of agentMessages) {
    if (msg.role === "assistant") {
      const u = (msg as AssistantMessage).usage;
      if (u) {
        inputTokens += u.input;
        outputTokens += u.output;
        cacheRead += u.cacheRead;
        cacheWrite += u.cacheWrite;
        totalCost += u.cost.total;
        lastInputTokens = u.input + u.cacheRead + u.cacheWrite;
      }
    }
  }
  return {
    inputTokens,
    outputTokens,
    cacheRead,
    cacheWrite,
    totalCost,
    lastInputTokens,
  };
}
