import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { StorageNamespace } from "../context";

// Minimal Model Context Protocol (MCP) client + tool loader.
// Speaks the Streamable HTTP transport (JSON-RPC over POST, SSE responses).
// Browser-only: uses fetch. For http:// servers behind mixed-content, point the
// url at a localhost proxy (Chromium treats http://localhost as trustworthy).

export interface McpServerConfig {
  name: string;
  url: string;
  enabled?: boolean;
  headers?: Record<string, string>;
}

export interface McpConfig {
  servers: McpServerConfig[];
}

interface McpToolDef {
  name: string;
  description?: string;
  inputSchema?: Record<string, unknown>;
}

interface McpContentBlock {
  type: string;
  text?: string;
  data?: string;
  mimeType?: string;
  [k: string]: unknown;
}

interface McpCallResult {
  content?: McpContentBlock[];
  structuredContent?: unknown;
  isError?: boolean;
}

const PROTOCOL_VERSION = "2025-06-18";

function mcpStorageKey(ns: StorageNamespace): string {
  return `${ns.localStoragePrefix}-mcp-config`;
}

export function loadMcpConfig(ns: StorageNamespace): McpConfig {
  try {
    const raw = localStorage.getItem(mcpStorageKey(ns));
    if (!raw) return { servers: [] };
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.servers)) return { servers: parsed.servers };
    return { servers: [] };
  } catch {
    return { servers: [] };
  }
}

export function saveMcpConfig(ns: StorageNamespace, config: McpConfig): void {
  localStorage.setItem(mcpStorageKey(ns), JSON.stringify(config));
}

function parseSseForId(text: string, id: number): unknown {
  let last: unknown;
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.startsWith("data:") ? line.slice(5).trim() : "";
    if (!trimmed) continue;
    try {
      const msg = JSON.parse(trimmed) as { id?: number };
      if (msg && msg.id === id) return msg;
      last = msg;
    } catch {
      // ignore non-JSON data lines (e.g. keep-alives)
    }
  }
  return last;
}

class McpClient {
  private url: string;
  private headers: Record<string, string>;
  private sessionId: string | null = null;
  private nextId = 1;

  constructor(url: string, headers?: Record<string, string>) {
    this.url = url;
    this.headers = headers ?? {};
  }

  private async rpc(
    method: string,
    params: Record<string, unknown> | undefined,
    notification: boolean,
  ): Promise<unknown> {
    const headers: Record<string, string> = {
      "content-type": "application/json",
      accept: "application/json, text/event-stream",
      ...this.headers,
    };
    if (this.sessionId) headers["mcp-session-id"] = this.sessionId;

    const id = notification ? undefined : this.nextId++;
    const body: Record<string, unknown> = { jsonrpc: "2.0", method };
    if (params !== undefined) body.params = params;
    if (id !== undefined) body.id = id;

    const res = await fetch(this.url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    const sid = res.headers.get("mcp-session-id");
    if (sid) this.sessionId = sid;

    if (notification) return undefined;
    if (!res.ok) {
      throw new Error(`MCP HTTP ${res.status}: ${await res.text()}`);
    }

    const contentType = res.headers.get("content-type") ?? "";
    let message: { result?: unknown; error?: { message?: string } } | undefined;
    if (contentType.includes("text/event-stream")) {
      message = parseSseForId(await res.text(), id as number) as typeof message;
    } else {
      message = (await res.json()) as typeof message;
    }
    if (message?.error) {
      throw new Error(message.error.message ?? "MCP error");
    }
    return message?.result;
  }

  async initialize(): Promise<void> {
    await this.rpc(
      "initialize",
      {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: {},
        clientInfo: { name: "office-agents", version: "1" },
      },
      false,
    );
    await this.rpc("notifications/initialized", undefined, true);
  }

  async listTools(): Promise<McpToolDef[]> {
    const result = (await this.rpc("tools/list", {}, false)) as {
      tools?: McpToolDef[];
    };
    return result?.tools ?? [];
  }

  async callTool(
    name: string,
    args: Record<string, unknown>,
  ): Promise<McpCallResult> {
    return (await this.rpc(
      "tools/call",
      { name, arguments: args },
      false,
    )) as McpCallResult;
  }
}

function sanitizeToolName(server: string, tool: string): string {
  const clean = `${server}_${tool}`.replace(/[^a-zA-Z0-9_-]/g, "_");
  return clean.slice(0, 64);
}

function wrapMcpTool(
  client: McpClient,
  server: McpServerConfig,
  tool: McpToolDef,
): AgentTool {
  const definition = {
    name: sanitizeToolName(server.name, tool.name),
    label: tool.name,
    description: tool.description ?? `MCP tool ${tool.name} (${server.name})`,
    parameters: tool.inputSchema ?? { type: "object", properties: {} },
    execute: async (_toolCallId: string, params: Record<string, unknown>) => {
      try {
        const result = await client.callTool(tool.name, params ?? {});
        const blocks = (result?.content ?? []).map((block) => {
          if (block.type === "text") {
            return { type: "text" as const, text: block.text ?? "" };
          }
          if (block.type === "image" && block.data) {
            return {
              type: "image" as const,
              data: block.data,
              mimeType: block.mimeType ?? "image/png",
            };
          }
          return { type: "text" as const, text: JSON.stringify(block) };
        });
        if (blocks.length === 0) {
          blocks.push({
            type: "text" as const,
            text: JSON.stringify(result?.structuredContent ?? { ok: true }),
          });
        }
        return { content: blocks, details: undefined };
      } catch (error) {
        const message =
          error instanceof Error ? error.message : "MCP tool call failed";
        return {
          content: [{ type: "text" as const, text: `Error: ${message}` }],
          details: undefined,
        };
      }
    },
  };
  return definition as unknown as AgentTool;
}

// Connect to every enabled MCP server, list its tools, and return them wrapped
// as AgentTools. A failing server is logged and skipped (never throws).
export async function loadMcpTools(ns: StorageNamespace): Promise<AgentTool[]> {
  const config = loadMcpConfig(ns);
  const tools: AgentTool[] = [];
  for (const server of config.servers) {
    if (server.enabled === false || !server.url) continue;
    try {
      const client = new McpClient(server.url, server.headers);
      await client.initialize();
      const mcpTools = await client.listTools();
      for (const tool of mcpTools) {
        tools.push(wrapMcpTool(client, server, tool));
      }
    } catch (error) {
      console.warn(`[mcp] server "${server.name}" failed to load:`, error);
    }
  }
  return tools;
}
