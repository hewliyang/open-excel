import {
  Agent,
  type AgentEvent,
  type AgentMessage,
  type ThinkingLevel as AgentThinkingLevel,
} from "@mariozechner/pi-agent-core";
import {
  type AssistantMessage,
  getModel,
  getModels,
  getProviders,
  type Model,
  streamSimple,
  type Usage,
} from "@mariozechner/pi-ai";
import type { ReactNode } from "react";
import { createContext, useCallback, useContext, useEffect, useRef, useState } from "react";
import { getWorkbookMetadata } from "../../../lib/excel/api";
import { EXCEL_TOOLS } from "../../../lib/tools";

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
    };

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  parts: MessagePart[];
  timestamp: number;
}

export type ThinkingLevel = "none" | "low" | "medium" | "high";

export interface ProviderConfig {
  provider: string;
  apiKey: string;
  model: string;
  useProxy: boolean;
  proxyUrl: string;
  thinking: ThinkingLevel;
}

export interface SessionStats {
  inputTokens: number;
  outputTokens: number;
  cacheRead: number;
  cacheWrite: number;
  totalCost: number;
  contextWindow: number;
  lastUsage: Usage | null;
}

const STORAGE_KEY = "openexcel-provider-config";

function loadSavedConfig(): ProviderConfig | null {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const config = JSON.parse(saved);
      if (config.proxyUrl === undefined) {
        config.proxyUrl = "";
      }
      return config;
    }
  } catch {}
  return null;
}

function applyProxyToModel(model: Model<any>, config: ProviderConfig): Model<any> {
  if (!config.useProxy || !config.proxyUrl || !model.baseUrl) return model;
  return {
    ...model,
    baseUrl: `${config.proxyUrl}/?url=${encodeURIComponent(model.baseUrl)}`,
  };
}

interface ChatState {
  messages: ChatMessage[];
  isStreaming: boolean;
  error: string | null;
  providerConfig: ProviderConfig | null;
  sessionStats: SessionStats;
}

const INITIAL_STATS: SessionStats = {
  inputTokens: 0,
  outputTokens: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalCost: 0,
  contextWindow: 0,
  lastUsage: null,
};

interface ChatContextValue {
  state: ChatState;
  sendMessage: (content: string) => Promise<void>;
  setProviderConfig: (config: ProviderConfig) => void;
  clearMessages: () => void;
  abort: () => void;
  availableProviders: string[];
  getModelsForProvider: (provider: string) => Model<any>[];
}

const ChatContext = createContext<ChatContextValue | null>(null);

const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:
READ:
- get_cell_ranges: Read cell values, formulas, and formatting
- get_range_as_csv: Get data as CSV (great for analysis)
- search_data: Find text across the spreadsheet
- get_all_objects: List charts, pivot tables, etc.

WRITE:
- set_cell_range: Write values, formulas, and formatting
- clear_cell_range: Clear contents or formatting
- copy_to: Copy ranges with formula translation
- modify_sheet_structure: Insert/delete/hide rows/columns, freeze panes
- modify_workbook_structure: Create/delete/rename sheets
- resize_range: Adjust column widths and row heights
- modify_object: Create/update/delete charts and pivot tables

Citations: Use markdown links with #cite: hash to reference sheets/cells. Clicking navigates there.
- Sheet only: [Sheet Name](#cite:sheetId)
- Cell/range: [A1:B10](#cite:sheetId!A1:B10)
Example: [Exchange Ratio](#cite:3) or [see cell B5](#cite:3!B5)

When the user asks about their data, read it first. Be concise. Use A1 notation for cell references.`;

function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function thinkingLevelToAgent(level: ThinkingLevel): AgentThinkingLevel {
  return level === "none" ? "off" : level;
}

function extractPartsFromAssistantMessage(message: AgentMessage, existingParts: MessagePart[] = []): MessagePart[] {
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

export function ChatProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<ChatState>(() => {
    const saved = loadSavedConfig();
    const validConfig = saved?.provider && saved?.apiKey && saved?.model ? saved : null;
    return {
      messages: [],
      isStreaming: false,
      error: null,
      providerConfig: validConfig,
      sessionStats: INITIAL_STATS,
    };
  });

  const agentRef = useRef<Agent | null>(null);
  const streamingMessageIdRef = useRef<string | null>(null);
  const isStreamingRef = useRef(false);
  const pendingConfigRef = useRef<ProviderConfig | null>(null);

  const availableProviders = getProviders();

  const getModelsForProvider = useCallback((provider: string): Model<any>[] => {
    try {
      return getModels(provider as any);
    } catch {
      return [];
    }
  }, []);

  const handleAgentEvent = useCallback((event: AgentEvent) => {
    switch (event.type) {
      case "message_start": {
        if (event.message.role === "assistant") {
          const id = generateId();
          streamingMessageIdRef.current = id;
          const parts = extractPartsFromAssistantMessage(event.message);
          const chatMessage: ChatMessage = {
            id,
            role: "assistant",
            parts,
            timestamp: event.message.timestamp,
          };
          setState((prev) => ({
            ...prev,
            messages: [...prev.messages, chatMessage],
          }));
        }
        break;
      }
      case "message_update": {
        if (event.message.role === "assistant" && streamingMessageIdRef.current) {
          setState((prev) => {
            const messages = [...prev.messages];
            const idx = messages.findIndex((m) => m.id === streamingMessageIdRef.current);
            if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(event.message, messages[idx].parts);
              messages[idx] = { ...messages[idx], parts };
            }
            return { ...prev, messages };
          });
        }
        break;
      }
      case "message_end": {
        if (event.message.role === "assistant") {
          const assistantMsg = event.message as AssistantMessage;
          setState((prev) => {
            const messages = [...prev.messages];
            const idx = messages.findIndex((m) => m.id === streamingMessageIdRef.current);
            if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(event.message, messages[idx].parts);
              messages[idx] = { ...messages[idx], parts };
            }
            console.log("[Chat] Assistant message result:", event.message);
            console.log("[Chat] Usage:", assistantMsg.usage);
            return {
              ...prev,
              messages,
              sessionStats: {
                inputTokens: prev.sessionStats.inputTokens + assistantMsg.usage.input,
                outputTokens: prev.sessionStats.outputTokens + assistantMsg.usage.output,
                cacheRead: prev.sessionStats.cacheRead + assistantMsg.usage.cacheRead,
                cacheWrite: prev.sessionStats.cacheWrite + assistantMsg.usage.cacheWrite,
                totalCost: prev.sessionStats.totalCost + assistantMsg.usage.cost.total,
                contextWindow: prev.sessionStats.contextWindow,
                lastUsage: assistantMsg.usage,
              },
            };
          });
          streamingMessageIdRef.current = null;
        }
        break;
      }
      case "tool_execution_start": {
        setState((prev) => {
          const messages = [...prev.messages];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex((p) => p.type === "toolCall" && p.id === event.toolCallId);
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                parts[partIdx] = { ...part, status: "running" };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return { ...prev, messages };
        });
        break;
      }
      case "tool_execution_update": {
        setState((prev) => {
          const messages = [...prev.messages];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex((p) => p.type === "toolCall" && p.id === event.toolCallId);
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                let partialText: string;
                if (typeof event.partialResult === "string") {
                  partialText = event.partialResult;
                } else if (event.partialResult?.content && Array.isArray(event.partialResult.content)) {
                  partialText = event.partialResult.content
                    .filter((c: { type: string }) => c.type === "text")
                    .map((c: { text: string }) => c.text)
                    .join("\n");
                } else {
                  partialText = JSON.stringify(event.partialResult, null, 2);
                }
                parts[partIdx] = { ...part, result: partialText };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return { ...prev, messages };
        });
        break;
      }
      case "tool_execution_end": {
        setState((prev) => {
          const messages = [...prev.messages];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex((p) => p.type === "toolCall" && p.id === event.toolCallId);
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                let resultText: string;
                if (typeof event.result === "string") {
                  resultText = event.result;
                } else if (event.result?.content && Array.isArray(event.result.content)) {
                  resultText = event.result.content
                    .filter((c: { type: string }) => c.type === "text")
                    .map((c: { text: string }) => c.text)
                    .join("\n");
                } else {
                  resultText = JSON.stringify(event.result, null, 2);
                }
                parts[partIdx] = { ...part, status: event.isError ? "error" : "complete", result: resultText };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return { ...prev, messages };
        });
        break;
      }
      case "agent_end": {
        isStreamingRef.current = false;
        setState((prev) => ({ ...prev, isStreaming: false }));
        streamingMessageIdRef.current = null;
        break;
      }
    }
  }, []);

  const applyConfig = useCallback(
    (config: ProviderConfig) => {
      let contextWindow = 0;
      let baseModel: Model<any>;
      try {
        baseModel = getModel(config.provider as any, config.model as any);
        contextWindow = baseModel.contextWindow;
      } catch {
        return;
      }

      const proxiedModel = applyProxyToModel(baseModel, config);
      const existingMessages = agentRef.current?.state.messages ?? [];

      if (agentRef.current) {
        agentRef.current.abort();
      }

      const agent = new Agent({
        initialState: {
          model: proxiedModel,
          systemPrompt: SYSTEM_PROMPT,
          thinkingLevel: thinkingLevelToAgent(config.thinking),
          tools: EXCEL_TOOLS,
          messages: existingMessages,
        },
        streamFn: (model, context, options) => {
          return streamSimple(model, context, {
            ...options,
            apiKey: config.apiKey,
          });
        },
      });
      agentRef.current = agent;
      agent.subscribe(handleAgentEvent);
      pendingConfigRef.current = null;

      console.log("[Chat] Model info:", {
        id: baseModel.id,
        contextWindow: baseModel.contextWindow,
        maxTokens: baseModel.maxTokens,
        cost: baseModel.cost,
        reasoning: baseModel.reasoning,
      });

      setState((prev) => ({
        ...prev,
        providerConfig: config,
        error: null,
        sessionStats: { ...prev.sessionStats, contextWindow },
      }));
    },
    [handleAgentEvent],
  );

  const setProviderConfig = useCallback(
    (config: ProviderConfig) => {
      if (isStreamingRef.current) {
        pendingConfigRef.current = config;
        setState((prev) => ({ ...prev, providerConfig: config }));
        return;
      }
      applyConfig(config);
    },
    [applyConfig],
  );

  const abort = useCallback(() => {
    agentRef.current?.abort();
    isStreamingRef.current = false;
    setState((prev) => ({ ...prev, isStreaming: false }));
  }, []);

  const sendMessage = useCallback(
    async (content: string) => {
      if (pendingConfigRef.current) {
        applyConfig(pendingConfigRef.current);
      }
      const agent = agentRef.current;
      if (!agent || !state.providerConfig) {
        setState((prev) => ({ ...prev, error: "Please configure your API key first" }));
        return;
      }

      const userMessage: ChatMessage = {
        id: generateId(),
        role: "user",
        parts: [{ type: "text", text: content }],
        timestamp: Date.now(),
      };

      isStreamingRef.current = true;
      setState((prev) => ({
        ...prev,
        messages: [...prev.messages, userMessage],
        isStreaming: true,
        error: null,
      }));

      try {
        let promptContent = content;
        try {
          console.log("[Chat] Fetching workbook metadata...");
          const metadata = await getWorkbookMetadata();
          console.log("[Chat] Workbook metadata:", metadata);
          promptContent = `<wb_context>\n${JSON.stringify(metadata, null, 2)}\n</wb_context>\n\n${content}`;
        } catch (err) {
          console.error("[Chat] Failed to get workbook metadata:", err);
        }
        await agent.prompt(promptContent);
        console.log("[Chat] Full context:", agent.state.messages);
      } catch (err) {
        isStreamingRef.current = false;
        setState((prev) => ({
          ...prev,
          isStreaming: false,
          error: err instanceof Error ? err.message : "An error occurred",
        }));
      }
    },
    [state.providerConfig, applyConfig],
  );

  const clearMessages = useCallback(() => {
    abort();
    agentRef.current?.reset();
    setState((prev) => ({ ...prev, messages: [], error: null, sessionStats: INITIAL_STATS }));
  }, [abort]);

  useEffect(() => {
    return () => {
      agentRef.current?.abort();
    };
  }, []);

  useEffect(() => {
    const saved = loadSavedConfig();
    if (saved?.provider && saved?.apiKey && saved?.model) {
      setProviderConfig(saved);
    }
  }, [setProviderConfig]);

  return (
    <ChatContext.Provider
      value={{
        state,
        sendMessage,
        setProviderConfig,
        clearMessages,
        abort,
        availableProviders,
        getModelsForProvider,
      }}
    >
      {children}
    </ChatContext.Provider>
  );
}

export function useChat() {
  const context = useContext(ChatContext);
  if (!context) throw new Error("useChat must be used within ChatProvider");
  return context;
}
