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
import type { DirtyRange } from "../../../lib/dirty-tracker";
import { getWorkbookMetadata, navigateTo } from "../../../lib/excel/api";
import {
  type ChatSession,
  createSession,
  deleteSession,
  getOrCreateCurrentSession,
  getOrCreateWorkbookId,
  getSession,
  listSessions,
  loadVfsFiles,
  saveSession,
  saveVfsFiles,
} from "../../../lib/storage";
import { EXCEL_TOOLS } from "../../../lib/tools";
import { deleteFile, listUploads, resetVfs, restoreVfs, snapshotVfs, writeFile } from "../../../lib/vfs";

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
  followMode: boolean;
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
      if (config.followMode === undefined) {
        config.followMode = true; // Default to on
      }
      return config;
    }
  } catch {}
  return null;
}

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

export interface UploadedFile {
  name: string;
  size: number;
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
  currentSession: ChatSession | null;
  sessions: ChatSession[];
  sheetNames: Record<number, string>;
  uploads: UploadedFile[];
  isUploading: boolean;
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
  sendMessage: (content: string, attachments?: string[]) => Promise<void>;
  setProviderConfig: (config: ProviderConfig) => void;
  clearMessages: () => void;
  abort: () => void;
  availableProviders: string[];
  getModelsForProvider: (provider: string) => Model<any>[];
  newSession: () => Promise<void>;
  switchSession: (sessionId: string) => Promise<void>;
  deleteCurrentSession: () => Promise<void>;
  getSheetName: (sheetId: number) => string | undefined;
  toggleFollowMode: () => void;
  processFiles: (files: File[]) => Promise<void>;
  removeUpload: (name: string) => Promise<void>;
}

const ChatContext = createContext<ChatContextValue | null>(null);

const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:

FILES & SHELL:
- read: Read uploaded files (images, CSV, text). Images are returned for visual analysis.
- bash: Execute bash commands in a sandboxed virtual filesystem. User uploads are in /home/user/uploads/.
  Supports: ls, cat, grep, find, awk, sed, jq, sort, uniq, wc, cut, head, tail, etc.

When the user uploads files, an <attachments> section lists their paths. Use read to access them.

EXCEL READ:
- get_cell_ranges: Read cell values, formulas, and formatting
- get_range_as_csv: Get data as CSV (great for analysis)
- search_data: Find text across the spreadsheet
- get_all_objects: List charts, pivot tables, etc.

EXCEL WRITE:
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
      currentSession: null,
      sessions: [],
      sheetNames: {},
      uploads: [],
      isUploading: false,
    };
  });

  const agentRef = useRef<Agent | null>(null);
  const streamingMessageIdRef = useRef<string | null>(null);
  const isStreamingRef = useRef(false);
  const pendingConfigRef = useRef<ProviderConfig | null>(null);
  const workbookIdRef = useRef<string | null>(null);
  const sessionLoadedRef = useRef(false);
  const currentSessionIdRef = useRef<string | null>(null);
  const followModeRef = useRef(state.providerConfig?.followMode ?? true);

  const availableProviders = getProviders();

  const getModelsForProvider = useCallback((provider: string): Model<any>[] => {
    try {
      return getModels(provider as any);
    } catch {
      return [];
    }
  }, []);

  const handleAgentEvent = useCallback((event: AgentEvent) => {
    console.log("[Chat] Agent event:", event.type, event);
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
          const isError = assistantMsg.stopReason === "error" || assistantMsg.stopReason === "aborted";
          console.log("[Chat] Assistant message result:", event.message);
          console.log("[Chat] Usage:", assistantMsg.usage);
          console.log("[Chat] stopReason:", assistantMsg.stopReason, "errorMessage:", assistantMsg.errorMessage);

          setState((prev) => {
            const messages = [...prev.messages];
            const idx = messages.findIndex((m) => m.id === streamingMessageIdRef.current);

            if (isError) {
              if (idx !== -1) {
                messages.splice(idx, 1);
              }
            } else if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(event.message, messages[idx].parts);
              messages[idx] = { ...messages[idx], parts };
            }

            return {
              ...prev,
              messages,
              error: isError ? assistantMsg.errorMessage || "Request failed" : prev.error,
              sessionStats: isError
                ? prev.sessionStats
                : {
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

        if (!event.isError && followModeRef.current) {
          const dirtyRanges = parseDirtyRanges(resultText);
          if (dirtyRanges && dirtyRanges.length > 0) {
            const first = dirtyRanges[0];
            if (first.sheetId >= 0 && first.range !== "*") {
              navigateTo(first.sheetId, first.range).catch((err) => {
                console.error("[FollowMode] Navigation failed:", err);
              });
            } else if (first.sheetId >= 0) {
              // For whole-sheet changes, just activate the sheet
              navigateTo(first.sheetId).catch((err) => {
                console.error("[FollowMode] Navigation failed:", err);
              });
            }
          }
        }

        setState((prev) => {
          const messages = [...prev.messages];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex((p) => p.type === "toolCall" && p.id === event.toolCallId);
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
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

      followModeRef.current = config.followMode ?? true;

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
    async (content: string, attachments?: string[]) => {
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

          if (metadata.sheetsMetadata) {
            const newSheetNames: Record<number, string> = {};
            for (const sheet of metadata.sheetsMetadata) {
              newSheetNames[sheet.id] = sheet.name;
            }
            setState((prev) => ({ ...prev, sheetNames: newSheetNames }));
          }
        } catch (err) {
          console.error("[Chat] Failed to get workbook metadata:", err);
        }

        // Add attachments section if files are uploaded
        if (attachments && attachments.length > 0) {
          const paths = attachments.map((name) => `/home/user/uploads/${name}`).join("\n");
          promptContent = `<attachments>\n${paths}\n</attachments>\n\n${promptContent}`;
        }

        await agent.prompt(promptContent);
        console.log("[Chat] Full context:", agent.state.messages);
      } catch (err) {
        console.error("[Chat] sendMessage error:", err);
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
    resetVfs();
    if (currentSessionIdRef.current) {
      Promise.all([saveSession(currentSessionIdRef.current, []), saveVfsFiles(currentSessionIdRef.current, [])]).catch(
        console.error,
      );
    }
    setState((prev) => ({ ...prev, messages: [], error: null, sessionStats: INITIAL_STATS, uploads: [] }));
  }, [abort]);

  const refreshSessions = useCallback(async () => {
    if (!workbookIdRef.current) return;
    const sessions = await listSessions(workbookIdRef.current);
    console.log(
      "[Chat] refreshSessions:",
      sessions.map((s) => ({ id: s.id, name: s.name, msgs: s.messages.length })),
    );
    setState((prev) => ({ ...prev, sessions }));
  }, []);

  const newSession = useCallback(async () => {
    console.log("[Chat] newSession called, workbookId:", workbookIdRef.current);
    if (!workbookIdRef.current) {
      console.error("[Chat] Cannot create session: workbookId not set");
      return;
    }
    if (isStreamingRef.current) {
      console.log("[Chat] newSession blocked: streaming in progress");
      return;
    }
    try {
      agentRef.current?.reset();
      resetVfs();
      const session = await createSession(workbookIdRef.current);
      console.log("[Chat] Created new session:", session.id);
      currentSessionIdRef.current = session.id;
      await refreshSessions();
      setState((prev) => ({
        ...prev,
        messages: [],
        currentSession: session,
        error: null,
        sessionStats: INITIAL_STATS,
        uploads: [],
      }));
    } catch (err) {
      console.error("[Chat] Failed to create session:", err);
    }
  }, [refreshSessions]);

  const switchSession = useCallback(async (sessionId: string) => {
    console.log("[Chat] switchSession called:", sessionId, "current:", currentSessionIdRef.current);
    if (currentSessionIdRef.current === sessionId) return;
    if (isStreamingRef.current) {
      console.log("[Chat] switchSession blocked: streaming in progress");
      return;
    }
    agentRef.current?.reset();
    try {
      const [session, vfsFiles] = await Promise.all([getSession(sessionId), loadVfsFiles(sessionId)]);
      console.log(
        "[Chat] Got session:",
        session?.id,
        "messages:",
        session?.messages.length,
        "vfs files:",
        vfsFiles.length,
      );
      if (!session) {
        console.error("[Chat] Session not found:", sessionId);
        return;
      }
      await restoreVfs(vfsFiles);
      currentSessionIdRef.current = session.id;
      const uploadNames = await listUploads();
      setState((prev) => ({
        ...prev,
        messages: session.messages,
        currentSession: session,
        error: null,
        sessionStats: INITIAL_STATS,
        uploads: uploadNames.map((name) => ({ name, size: 0 })),
      }));
    } catch (err) {
      console.error("[Chat] Failed to switch session:", err);
    }
  }, []);

  const deleteCurrentSession = useCallback(async () => {
    if (!currentSessionIdRef.current || !workbookIdRef.current) return;
    if (isStreamingRef.current) {
      console.log("[Chat] deleteCurrentSession blocked: streaming in progress");
      return;
    }
    agentRef.current?.reset();
    const deletedId = currentSessionIdRef.current;
    await Promise.all([deleteSession(deletedId), saveVfsFiles(deletedId, [])]);
    const session = await getOrCreateCurrentSession(workbookIdRef.current);
    currentSessionIdRef.current = session.id;
    // Restore VFS for the session we're switching to
    const vfsFiles = await loadVfsFiles(session.id);
    await restoreVfs(vfsFiles);
    await refreshSessions();
    const uploadNames = await listUploads();
    setState((prev) => ({
      ...prev,
      messages: session.messages,
      currentSession: session,
      error: null,
      sessionStats: INITIAL_STATS,
      uploads: uploadNames.map((name) => ({ name, size: 0 })),
    }));
  }, [refreshSessions]);

  const prevStreamingRef = useRef(false);
  useEffect(() => {
    if (prevStreamingRef.current && !state.isStreaming && currentSessionIdRef.current) {
      const sessionId = currentSessionIdRef.current;
      // Snapshot VFS first (returns native Promise), then pass result to Dexie.
      // Never nest Dexie calls inside non-Dexie .then() — breaks PSD zone tracking in Office SES sandbox.
      (async () => {
        try {
          const vfsFiles = await snapshotVfs();
          await Promise.all([saveSession(sessionId, state.messages), saveVfsFiles(sessionId, vfsFiles)]);
          await refreshSessions();
          const updated = await getSession(sessionId);
          if (updated) {
            setState((prev) => ({ ...prev, currentSession: updated }));
          }
        } catch (e) {
          console.error(e);
        }
      })();
    }
    prevStreamingRef.current = state.isStreaming;
  }, [state.isStreaming, state.messages, refreshSessions]);

  useEffect(() => {
    return () => {
      agentRef.current?.abort();
    };
  }, []);

  useEffect(() => {
    if (sessionLoadedRef.current) return;
    sessionLoadedRef.current = true;

    const saved = loadSavedConfig();
    if (saved?.provider && saved?.apiKey && saved?.model) {
      applyConfig(saved);
    }

    getOrCreateWorkbookId()
      .then(async (id) => {
        workbookIdRef.current = id;
        console.log("[Chat] Workbook ID:", id);
        const session = await getOrCreateCurrentSession(id);
        currentSessionIdRef.current = session.id;
        const [sessions, vfsFiles] = await Promise.all([listSessions(id), loadVfsFiles(session.id)]);
        if (vfsFiles.length > 0) {
          await restoreVfs(vfsFiles);
        }
        const uploadNames = await listUploads();
        console.log(
          "[Chat] Loaded session:",
          session.id,
          "with",
          session.messages.length,
          "messages,",
          vfsFiles.length,
          "vfs files",
        );
        setState((prev) => ({
          ...prev,
          messages: session.messages,
          currentSession: session,
          sessions,
          uploads: uploadNames.map((name) => ({ name, size: 0 })),
        }));
      })
      .catch((err) => {
        console.error("[Chat] Failed to load session:", err);
      });
  }, [applyConfig]);

  const getSheetName = useCallback(
    (sheetId: number): string | undefined => state.sheetNames[sheetId],
    [state.sheetNames],
  );

  const processFiles = useCallback(async (files: File[]) => {
    if (files.length === 0) return;
    setState((prev) => ({ ...prev, isUploading: true }));
    try {
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        await writeFile(file.name, data);
        setState((prev) => {
          const exists = prev.uploads.some((u) => u.name === file.name);
          if (exists) {
            return {
              ...prev,
              uploads: prev.uploads.map((u) => (u.name === file.name ? { name: file.name, size: file.size } : u)),
            };
          }
          return { ...prev, uploads: [...prev.uploads, { name: file.name, size: file.size }] };
        });
      }
      if (currentSessionIdRef.current) {
        const snapshot = await snapshotVfs();
        await saveVfsFiles(currentSessionIdRef.current, snapshot);
      }
    } catch (err) {
      console.error("Failed to upload file:", err);
    } finally {
      setState((prev) => ({ ...prev, isUploading: false }));
    }
  }, []);

  const removeUpload = useCallback(async (name: string) => {
    try {
      await deleteFile(name);
      setState((prev) => ({ ...prev, uploads: prev.uploads.filter((u) => u.name !== name) }));
      if (currentSessionIdRef.current) {
        const snapshot = await snapshotVfs();
        await saveVfsFiles(currentSessionIdRef.current, snapshot);
      }
    } catch (err) {
      console.error("Failed to delete file:", err);
      setState((prev) => ({ ...prev, uploads: prev.uploads.filter((u) => u.name !== name) }));
    }
  }, []);

  const toggleFollowMode = useCallback(() => {
    setState((prev) => {
      if (!prev.providerConfig) return prev;
      const newFollowMode = !prev.providerConfig.followMode;
      followModeRef.current = newFollowMode;
      const newConfig = { ...prev.providerConfig, followMode: newFollowMode };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(newConfig));
      return { ...prev, providerConfig: newConfig };
    });
  }, []);

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
        newSession,
        switchSession,
        deleteCurrentSession,
        getSheetName,
        toggleFollowMode,
        processFiles,
        removeUpload,
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
