import type { Api, Model } from "@earendil-works/pi-ai";
import {
  type AgentContext,
  AgentRuntime,
  type ProviderConfig,
  type RuntimeState,
} from "@office-agents/sdk";
import { get, type Writable, writable } from "svelte/store";
import type { AppAdapter } from "./app-adapter";

export class ChatController {
  readonly context: AgentContext;
  readonly state: Writable<RuntimeState>;
  adapter: AppAdapter;
  #runtime: AgentRuntime;
  #unsubscribe: (() => void) | null = null;

  constructor(adapter: AppAdapter, context: AgentContext) {
    this.adapter = adapter;
    this.context = context;
    this.#runtime = new AgentRuntime(adapter, this.context);
    this.state = writable(this.#runtime.getState());
    this.#unsubscribe = this.#runtime.subscribe((next) => this.state.set(next));
    this.#runtime.init();
  }

  get snapshot() {
    return get(this.state);
  }

  get availableProviders() {
    return this.#runtime.getAvailableProviders();
  }

  setAdapter(adapter: AppAdapter) {
    this.adapter = adapter;
    this.#runtime.setAdapter(adapter);
  }

  dispose() {
    this.#unsubscribe?.();
    this.#unsubscribe = null;
    this.#runtime.dispose();
  }

  getModelsForProvider(provider: string): Model<Api>[] {
    return this.#runtime.getModelsForProvider(provider);
  }

  sendMessage(content: string, attachments?: string[]) {
    return this.#runtime.sendMessage(content, attachments);
  }

  setProviderConfig(config: ProviderConfig) {
    this.#runtime.setProviderConfig(config);
  }

  reloadMcpTools() {
    return this.#runtime.reloadMcpTools();
  }

  clearMessages() {
    this.#runtime.clearMessages();
  }

  abort() {
    this.#runtime.abort();
  }

  newSession() {
    return this.#runtime.newSession();
  }

  switchSession(sessionId: string) {
    return this.#runtime.switchSession(sessionId);
  }

  deleteCurrentSession() {
    return this.#runtime.deleteCurrentSession();
  }

  getName(id: number) {
    return this.#runtime.getName(id);
  }

  toggleFollowMode() {
    this.#runtime.toggleFollowMode();
  }

  toggleExpandToolCalls() {
    this.#runtime.toggleExpandToolCalls();
  }

  async processFiles(files: File[]) {
    if (files.length === 0) return;

    const inputs = await Promise.all(
      files.map(async (file) => ({
        name: file.name,
        size: file.size,
        data: new Uint8Array(await file.arrayBuffer()),
      })),
    );
    await this.#runtime.uploadFiles(inputs);
  }

  removeUpload(name: string) {
    return this.#runtime.removeUpload(name);
  }

  async installSkill(files: File[]) {
    if (files.length === 0) return;

    const inputs = await Promise.all(
      files.map(async (file) => {
        const fullPath = file.webkitRelativePath || file.name;
        const parts = fullPath.split("/");
        const path = parts.length > 1 ? parts.slice(1).join("/") : parts[0];
        return { path, data: new Uint8Array(await file.arrayBuffer()) };
      }),
    );

    await this.#runtime.installSkill(inputs);
  }

  uninstallSkill(name: string) {
    return this.#runtime.uninstallSkill(name);
  }
}
