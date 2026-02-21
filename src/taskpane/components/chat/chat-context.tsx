import {
  Agent,
  type AgentEvent,
  type ThinkingLevel as AgentThinkingLevel,
} from "@mariozechner/pi-agent-core";
import {
  type AssistantMessage,
  getModel,
  getModels,
  getProviders,
  type Model,
  streamSimple,
} from "@mariozechner/pi-ai";
import type { ReactNode } from "react";
import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useRef,
  useState,
} from "react";
import type { DirtyRange } from "../../../lib/dirty-tracker";
import { getWorkbookMetadata, navigateTo } from "../../../lib/excel/api";
import {
  agentMessagesToChatMessages,
  type ChatMessage,
  deriveStats,
  extractPartsFromAssistantMessage,
  generateId,
  type SessionStats,
} from "../../../lib/message-utils";
import {
  loadOAuthCredentials,
  refreshOAuthToken,
  saveOAuthCredentials,
} from "../../../lib/oauth";
import {
  applyProxyToModel,
  buildCustomModel,
  loadSavedConfig,
  type ProviderConfig,
  saveConfig,
  type ThinkingLevel,
} from "../../../lib/provider-config";
import {
  addSkill,
  buildSkillsPromptSection,
  getInstalledSkills,
  removeSkill,
  type SkillMeta,
  syncSkillsToVfs,
} from "../../../lib/skills";
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
import {
  deleteFile,
  listUploads,
  resetVfs,
  restoreVfs,
  snapshotVfs,
  writeFile,
} from "../../../lib/vfs";

export type {
  ChatMessage,
  MessagePart,
  SessionStats,
  ToolCallStatus,
} from "../../../lib/message-utils";
export type { ProviderConfig, ThinkingLevel };

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
  skills: SkillMeta[];
}

const INITIAL_STATS: SessionStats = { ...deriveStats([]), contextWindow: 0 };

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
  installSkill: (files: File[]) => Promise<void>;
  uninstallSkill: (name: string) => Promise<void>;
}

const ChatContext = createContext<ChatContextValue | null>(null);

function buildSystemPrompt(skills: SkillMeta[]): string {
  return `You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:

FILES & SHELL:
- read: Read uploaded files (images, CSV, text). Images are returned for visual analysis.
- bash: Execute bash commands in a sandboxed virtual filesystem. User uploads are in /home/user/uploads/.
  Supports: ls, cat, grep, find, awk, sed, jq, sort, uniq, wc, cut, head, tail, etc.

  Custom commands for efficient data transfer (data flows directly, never enters your context):
  - csv-to-sheet <file> <sheetId> [startCell] [--force] — Import CSV from VFS into spreadsheet. Auto-detects types.
    Fails if target cells already have data. Use --force to overwrite (confirm with user first).
  - sheet-to-csv <sheetId> [range] [file] — Export range to CSV. Defaults to full used range if no range given. Prints to stdout if no file given (pipeable).
  - pdf-to-text <file> <outfile> — Extract text from PDF to file. Use head/grep/tail to read selectively.
  - pdf-to-images <file> <outdir> [--scale=N] [--pages=1,3,5-8] — Render PDF pages to PNG images. Use for scanned PDFs where text extraction won't work. Then use read to visually inspect the images.
  - docx-to-text <file> <outfile> — Extract text from DOCX to file.
  - xlsx-to-csv <file> <outfile> [sheet] — Convert XLSX/XLS/ODS sheet to CSV. Sheet by name or 0-based index.
  - image-to-sheet <file> <width> <height> <sheetId> [startCell] [--cell-size=N] — Render an image as pixel art in Excel. Downsamples to target size and paints cell backgrounds. Cell size in points (default: 3). Max 500×500. Example: image-to-sheet uploads/logo.png 64 64 1 A1 --cell-size=4
  - web-search <query> [--max=N] [--region=REGION] [--time=d|w|m|y] [--page=N] [--json] — Search the web. Returns title, URL, and snippet for each result.
  - web-fetch <url> <outfile> — Fetch a web page and extract its readable content to a file. Use head/grep/tail to read selectively.

  Examples:
    csv-to-sheet uploads/data.csv 1 A1       # import CSV to sheet 1
    sheet-to-csv 1 export.csv                 # export entire sheet to file
    sheet-to-csv 1 A1:D100 export.csv         # export specific range to file
    sheet-to-csv 1 | sort -t, -k3 -rn | head -20   # pipe entire sheet to analysis
    cut -d, -f1,3 uploads/data.csv > filtered.csv && csv-to-sheet filtered.csv 1 A1  # filter then import
    web-search "S&P 500 companies list"       # search the web
    web-search "USD EUR exchange rate" --max=5 --time=w  # recent results only
    web-fetch https://example.com/article page.txt && grep -i "revenue" page.txt  # fetch then grep

  IMPORTANT: When importing file data into the spreadsheet, ALWAYS prefer csv-to-sheet over reading
  the file content and calling set_cell_range. This avoids wasting tokens on data that doesn't need
  to pass through your context.

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

When the user asks about their data, read it first. Be concise. Use A1 notation for cell references.

${buildSkillsPromptSection(skills)}
`;
}

function thinkingLevelToAgent(level: ThinkingLevel): AgentThinkingLevel {
  return level === "none" ? "off" : level;
}

export function ChatProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<ChatState>(() => {
    const saved = loadSavedConfig();
    const validConfig =
      saved?.provider && saved?.apiKey && saved?.model ? saved : null;
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
      skills: [],
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
  const skillsRef = useRef<SkillMeta[]>([]);

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
        if (
          event.message.role === "assistant" &&
          streamingMessageIdRef.current
        ) {
          setState((prev) => {
            const messages = [...prev.messages];
            const idx = messages.findIndex(
              (m) => m.id === streamingMessageIdRef.current,
            );
            if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(
                event.message,
                messages[idx].parts,
              );
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
          const isError =
            assistantMsg.stopReason === "error" ||
            assistantMsg.stopReason === "aborted";
          console.log("[Chat] Assistant message result:", event.message);
          console.log("[Chat] Usage:", assistantMsg.usage);
          console.log(
            "[Chat] stopReason:",
            assistantMsg.stopReason,
            "errorMessage:",
            assistantMsg.errorMessage,
          );

          setState((prev) => {
            const messages = [...prev.messages];
            const idx = messages.findIndex(
              (m) => m.id === streamingMessageIdRef.current,
            );

            if (isError) {
              if (idx !== -1) {
                messages.splice(idx, 1);
              }
            } else if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(
                event.message,
                messages[idx].parts,
              );
              messages[idx] = { ...messages[idx], parts };
            }

            return {
              ...prev,
              messages,
              error: isError
                ? assistantMsg.errorMessage || "Request failed"
                : prev.error,
              sessionStats: isError
                ? prev.sessionStats
                : {
                    ...deriveStats(agentRef.current?.state.messages ?? []),
                    contextWindow: prev.sessionStats.contextWindow,
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
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
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
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                let partialText: string;
                if (typeof event.partialResult === "string") {
                  partialText = event.partialResult;
                } else if (
                  event.partialResult?.content &&
                  Array.isArray(event.partialResult.content)
                ) {
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
        let resultImages: { data: string; mimeType: string }[] | undefined;
        if (typeof event.result === "string") {
          resultText = event.result;
        } else if (
          event.result?.content &&
          Array.isArray(event.result.content)
        ) {
          resultText = event.result.content
            .filter((c: { type: string }) => c.type === "text")
            .map((c: { text: string }) => c.text)
            .join("\n");
          const images = event.result.content
            .filter((c: { type: string }) => c.type === "image")
            .map((c: { data: string; mimeType: string }) => ({
              data: c.data,
              mimeType: c.mimeType,
            }));
          if (images.length > 0) resultImages = images;
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
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                parts[partIdx] = {
                  ...part,
                  status: event.isError ? "error" : "complete",
                  result: resultText,
                  images: resultImages,
                };
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

  const configRef = useRef<ProviderConfig | null>(null);

  const getActiveApiKey = useCallback(
    async (config: ProviderConfig): Promise<string> => {
      if (config.authMethod !== "oauth") {
        return config.apiKey;
      }
      const creds = loadOAuthCredentials(config.provider);
      if (!creds) return config.apiKey;
      if (Date.now() < creds.expires) {
        return creds.access;
      }
      console.log("[Chat] Refreshing OAuth token before API call...");
      const refreshed = await refreshOAuthToken(
        config.provider,
        creds.refresh,
        config.proxyUrl,
        config.useProxy,
      );
      saveOAuthCredentials(config.provider, refreshed);
      console.log("[Chat] OAuth token refreshed");
      return refreshed.access;
    },
    [],
  );

  const applyConfig = useCallback(
    (config: ProviderConfig) => {
      let contextWindow = 0;
      let baseModel: Model<any>;
      if (config.provider === "custom") {
        const custom = buildCustomModel(config);
        if (!custom) return;
        baseModel = custom;
      } else {
        try {
          baseModel = getModel(config.provider as any, config.model as any);
        } catch {
          return;
        }
      }
      contextWindow = baseModel.contextWindow;
      configRef.current = config;

      const proxiedModel = applyProxyToModel(baseModel, config);
      const existingMessages = agentRef.current?.state.messages ?? [];

      if (agentRef.current) {
        agentRef.current.abort();
      }

      const systemPrompt = buildSystemPrompt(skillsRef.current);
      console.log(
        "[Chat] Skills in prompt:",
        skillsRef.current.length,
        skillsRef.current.map((s) => s.name),
      );
      console.log("[Chat] System prompt tail:", systemPrompt.slice(-500));

      const agent = new Agent({
        initialState: {
          model: proxiedModel,
          systemPrompt,
          thinkingLevel: thinkingLevelToAgent(config.thinking),
          tools: EXCEL_TOOLS,
          messages: existingMessages,
        },
        streamFn: async (model, context, options) => {
          const cfg = configRef.current ?? config;
          const apiKey = await getActiveApiKey(cfg);
          return streamSimple(model, context, {
            ...options,
            apiKey,
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
    [handleAgentEvent, getActiveApiKey],
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
        setState((prev) => ({
          ...prev,
          error: "Please configure your API key first",
        }));
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
          const paths = attachments
            .map((name) => `/home/user/uploads/${name}`)
            .join("\n");
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
      Promise.all([
        saveSession(currentSessionIdRef.current, []),
        saveVfsFiles(currentSessionIdRef.current, []),
      ]).catch(console.error);
    }
    setState((prev) => ({
      ...prev,
      messages: [],
      error: null,
      sessionStats: INITIAL_STATS,
      uploads: [],
    }));
  }, [abort]);

  const refreshSessions = useCallback(async () => {
    if (!workbookIdRef.current) return;
    const sessions = await listSessions(workbookIdRef.current);
    console.log(
      "[Chat] refreshSessions:",
      sessions.map((s) => ({
        id: s.id,
        name: s.name,
        msgs: (s.agentMessages ?? []).length,
      })),
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
    console.log(
      "[Chat] switchSession called:",
      sessionId,
      "current:",
      currentSessionIdRef.current,
    );
    if (currentSessionIdRef.current === sessionId) return;
    if (isStreamingRef.current) {
      console.log("[Chat] switchSession blocked: streaming in progress");
      return;
    }
    agentRef.current?.reset();
    try {
      const [session, vfsFiles] = await Promise.all([
        getSession(sessionId),
        loadVfsFiles(sessionId),
      ]);
      console.log(
        "[Chat] switchSession loaded:",
        session?.id,
        "agentMessages:",
        session?.agentMessages.length,
        "vfs:",
        vfsFiles.length,
      );
      if (!session) {
        console.error("[Chat] Session not found:", sessionId);
        return;
      }
      await restoreVfs(vfsFiles);
      currentSessionIdRef.current = session.id;

      if (session.agentMessages.length > 0 && agentRef.current) {
        agentRef.current.replaceMessages(session.agentMessages);
      }

      const uploadNames = await listUploads();
      const stats = deriveStats(session.agentMessages);
      setState((prev) => ({
        ...prev,
        messages: agentMessagesToChatMessages(session.agentMessages),
        currentSession: session,
        error: null,
        sessionStats: {
          ...stats,
          contextWindow: prev.sessionStats.contextWindow,
        },
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
    const vfsFiles = await loadVfsFiles(session.id);
    await restoreVfs(vfsFiles);

    if (session.agentMessages.length > 0 && agentRef.current) {
      agentRef.current.replaceMessages(session.agentMessages);
    }

    await refreshSessions();
    const uploadNames = await listUploads();
    const stats = deriveStats(session.agentMessages);
    setState((prev) => ({
      ...prev,
      messages: agentMessagesToChatMessages(session.agentMessages),
      currentSession: session,
      error: null,
      sessionStats: {
        ...stats,
        contextWindow: prev.sessionStats.contextWindow,
      },
      uploads: uploadNames.map((name) => ({ name, size: 0 })),
    }));
  }, [refreshSessions]);

  const prevStreamingRef = useRef(false);
  useEffect(() => {
    if (
      prevStreamingRef.current &&
      !state.isStreaming &&
      currentSessionIdRef.current
    ) {
      const sessionId = currentSessionIdRef.current;
      const agentMessages = agentRef.current?.state.messages ?? [];
      // Snapshot VFS first (returns native Promise), then save to IndexedDB.
      (async () => {
        try {
          const vfsFiles = await snapshotVfs();
          await Promise.all([
            saveSession(sessionId, agentMessages),
            saveVfsFiles(sessionId, vfsFiles),
          ]);
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
  }, [state.isStreaming, refreshSessions]);

  useEffect(() => {
    return () => {
      agentRef.current?.abort();
    };
  }, []);

  useEffect(() => {
    if (sessionLoadedRef.current) return;
    sessionLoadedRef.current = true;

    getOrCreateWorkbookId()
      .then(async (id) => {
        workbookIdRef.current = id;
        console.log("[Chat] Workbook ID:", id);

        // Load skills into VFS cache BEFORE applyConfig so the system prompt includes them
        const skills = await getInstalledSkills();
        skillsRef.current = skills;
        await syncSkillsToVfs();
        console.log(
          "[Chat] Loaded",
          skills.length,
          "skills:",
          skills.map((s) => s.name),
        );

        // Now apply provider config — agent gets the correct system prompt with skills
        const saved = loadSavedConfig();
        if (saved?.provider && saved?.apiKey && saved?.model) {
          applyConfig(saved);
        }

        const session = await getOrCreateCurrentSession(id);
        currentSessionIdRef.current = session.id;
        const [sessions, vfsFiles] = await Promise.all([
          listSessions(id),
          loadVfsFiles(session.id),
        ]);
        if (vfsFiles.length > 0) {
          await restoreVfs(vfsFiles);
        }

        if (session.agentMessages.length > 0 && agentRef.current) {
          agentRef.current.replaceMessages(session.agentMessages);
          console.log(
            "[Chat] Restored",
            session.agentMessages.length,
            "agent messages from DB",
          );
        }

        const uploadNames = await listUploads();
        const stats = deriveStats(session.agentMessages);
        console.log(
          "[Chat] Loaded session:",
          session.id,
          "agentMessages:",
          session.agentMessages.length,
          "vfs:",
          vfsFiles.length,
        );
        setState((prev) => ({
          ...prev,
          messages: agentMessagesToChatMessages(session.agentMessages),
          currentSession: session,
          sessions,
          skills,
          sessionStats: {
            ...stats,
            contextWindow: prev.sessionStats.contextWindow,
          },
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
              uploads: prev.uploads.map((u) =>
                u.name === file.name ? { name: file.name, size: file.size } : u,
              ),
            };
          }
          return {
            ...prev,
            uploads: [...prev.uploads, { name: file.name, size: file.size }],
          };
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
      setState((prev) => ({
        ...prev,
        uploads: prev.uploads.filter((u) => u.name !== name),
      }));
      if (currentSessionIdRef.current) {
        const snapshot = await snapshotVfs();
        await saveVfsFiles(currentSessionIdRef.current, snapshot);
      }
    } catch (err) {
      console.error("Failed to delete file:", err);
      setState((prev) => ({
        ...prev,
        uploads: prev.uploads.filter((u) => u.name !== name),
      }));
    }
  }, []);

  const refreshSkillsAndRebuildAgent = useCallback(async () => {
    skillsRef.current = await getInstalledSkills();
    setState((prev) => {
      // Re-apply config to rebuild agent with updated system prompt
      if (prev.providerConfig) {
        applyConfig(prev.providerConfig);
      }
      return { ...prev, skills: skillsRef.current };
    });
  }, [applyConfig]);

  const installSkill = useCallback(
    async (files: File[]) => {
      if (files.length === 0) return;
      try {
        const inputs = await Promise.all(
          files.map(async (f) => {
            // For folder uploads, webkitRelativePath is "folderName/file.md"
            // Strip the top-level folder to get the relative path within the skill
            const fullPath = f.webkitRelativePath || f.name;
            const parts = fullPath.split("/");
            const path = parts.length > 1 ? parts.slice(1).join("/") : parts[0];
            return { path, data: new Uint8Array(await f.arrayBuffer()) };
          }),
        );
        const meta = await addSkill(inputs);
        console.log("[Chat] Installed skill:", meta.name);
        await refreshSkillsAndRebuildAgent();
      } catch (err) {
        console.error("[Chat] Failed to install skill:", err);
        setState((prev) => ({
          ...prev,
          error: err instanceof Error ? err.message : "Failed to install skill",
        }));
      }
    },
    [refreshSkillsAndRebuildAgent],
  );

  const uninstallSkill = useCallback(
    async (name: string) => {
      try {
        await removeSkill(name);
        console.log("[Chat] Uninstalled skill:", name);
        await refreshSkillsAndRebuildAgent();
      } catch (err) {
        console.error("[Chat] Failed to uninstall skill:", err);
      }
    },
    [refreshSkillsAndRebuildAgent],
  );

  const toggleFollowMode = useCallback(() => {
    setState((prev) => {
      if (!prev.providerConfig) return prev;
      const newFollowMode = !prev.providerConfig.followMode;
      followModeRef.current = newFollowMode;
      const newConfig = { ...prev.providerConfig, followMode: newFollowMode };
      saveConfig(newConfig);
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
        installSkill,
        uninstallSkill,
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
