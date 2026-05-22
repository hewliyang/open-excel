import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { UserMessage } from "@earendil-works/pi-ai";
import { type DBSchema, type IDBPDatabase, openDB } from "idb";
import type { StorageNamespace } from "../context";
import { stripEnrichment } from "../message-utils";

export interface ChatSession {
  id: string;
  workbookId: string;
  name: string;
  agentMessages: AgentMessage[];
  createdAt: number;
  updatedAt: number;
}

export interface VfsFile {
  id: string;
  sessionId: string;
  path: string;
  data: Uint8Array;
}

export interface SkillFile {
  id: string;
  skillName: string;
  path: string;
  data: Uint8Array;
}

interface OfficeAgentsSchema extends DBSchema {
  sessions: {
    key: string;
    value: ChatSession;
    indexes: { workbookId: string; updatedAt: number };
  };
  vfsFiles: {
    key: string;
    value: VfsFile;
    indexes: { sessionId: string };
  };
  skillFiles: {
    key: string;
    value: SkillFile;
    indexes: { skillName: string };
  };
}

let dbPromise: Promise<IDBPDatabase<OfficeAgentsSchema>> | null = null;
let dbKey: string | null = null;
let useMemoryFallback = false;

// NOTE: In-memory fallback for when IndexedDB is blocked (e.g., PowerPoint Online)
const memoryStore: {
  sessions: Map<string, ChatSession>;
  vfsFiles: Map<string, VfsFile>;
  skillFiles: Map<string, SkillFile>;
} = {
  sessions: new Map(),
  vfsFiles: new Map(),
  skillFiles: new Map(),
};

async function getDb(
  ns: StorageNamespace,
): Promise<IDBPDatabase<OfficeAgentsSchema> | null> {
  if (useMemoryFallback) return null;

  const key = `${ns.dbName}@${ns.dbVersion}`;
  if (dbPromise && dbKey === key) return dbPromise;

  dbKey = key;
  try {
    dbPromise = openDB<OfficeAgentsSchema>(ns.dbName, ns.dbVersion, {
      upgrade(db) {
        if (!db.objectStoreNames.contains("sessions")) {
          const sessions = db.createObjectStore("sessions", { keyPath: "id" });
          sessions.createIndex("workbookId", "workbookId");
          sessions.createIndex("updatedAt", "updatedAt");
        }
        if (!db.objectStoreNames.contains("vfsFiles")) {
          const vfsFiles = db.createObjectStore("vfsFiles", { keyPath: "id" });
          vfsFiles.createIndex("sessionId", "sessionId");
        }
        if (!db.objectStoreNames.contains("skillFiles")) {
          const skillFiles = db.createObjectStore("skillFiles", {
            keyPath: "id",
          });
          skillFiles.createIndex("skillName", "skillName");
        }
      },
    });
    return await dbPromise;
  } catch (error) {
    // NOTE: IndexedDB blocked (PowerPoint Online with tracking protection)
    console.warn("[DB] IndexedDB unavailable, using in-memory fallback:", error);
    useMemoryFallback = true;
    dbPromise = null;
    return null;
  }
}

function extractUserText(msg: AgentMessage): string | null {
  if (msg.role !== "user") return null;
  const text = stripEnrichment((msg as UserMessage).content).trim();
  return text || null;
}

function deriveSessionName(agentMessages: AgentMessage[]): string | null {
  const firstUser = agentMessages.find((m) => m.role === "user");
  if (!firstUser) return null;
  const text = extractUserText(firstUser);
  if (!text) return null;
  return text.length > 40 ? `${text.slice(0, 37)}...` : text;
}

export function getSessionMessageCount(session: ChatSession): number {
  return (session.agentMessages ?? []).filter(
    (m) => m.role === "user" || m.role === "assistant",
  ).length;
}

export async function getOrCreateDocumentId(
  ns: StorageNamespace,
): Promise<string> {
  const settingsKey =
    ns.documentIdSettingsKey ?? `${ns.documentSettingsPrefix}-document-id`;
  return new Promise((resolve, reject) => {
    const settings = Office.context.document.settings;
    let docId = settings.get(settingsKey) as string | null;

    if (docId) {
      resolve(docId);
      return;
    }

    docId = crypto.randomUUID();
    settings.set(settingsKey, docId);
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(docId);
      } else {
        reject(
          new Error(result.error?.message ?? "Failed to save document ID"),
        );
      }
    });
  });
}

export async function listSessions(
  ns: StorageNamespace,
  workbookId: string,
): Promise<ChatSession[]> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const sessions = [...memoryStore.sessions.values()]
      .filter(s => s.workbookId === workbookId);
    for (const s of sessions) {
      if (!s.agentMessages) s.agentMessages = [];
    }
    sessions.sort((a, b) => b.updatedAt - a.updatedAt);
    return sessions;
  }

  const sessions = await db.getAllFromIndex(
    "sessions",
    "workbookId",
    workbookId,
  );
  for (const s of sessions) {
    if (!s.agentMessages) s.agentMessages = [];
  }
  sessions.sort((a, b) => b.updatedAt - a.updatedAt);
  return sessions;
}

export async function createSession(
  ns: StorageNamespace,
  workbookId: string,
  name?: string,
): Promise<ChatSession> {
  const db = await getDb(ns);
  const now = Date.now();
  const session: ChatSession = {
    id: crypto.randomUUID(),
    workbookId,
    name: name ?? "New Chat",
    agentMessages: [],
    createdAt: now,
    updatedAt: now,
  };

  // NOTE: Memory fallback
  if (!db) {
    memoryStore.sessions.set(session.id, session);
    return session;
  }

  await db.add("sessions", session);
  return session;
}

export async function getSession(
  ns: StorageNamespace,
  sessionId: string,
): Promise<ChatSession | undefined> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const session = memoryStore.sessions.get(sessionId);
    if (session && !session.agentMessages) {
      session.agentMessages = [];
    }
    return session;
  }

  const session = await db.get("sessions", sessionId);
  if (session && !session.agentMessages) {
    session.agentMessages = [];
  }
  return session;
}

export async function saveSession(
  ns: StorageNamespace,
  sessionId: string,
  agentMessages: AgentMessage[],
): Promise<void> {
  console.log(
    "[DB] saveSession:",
    sessionId,
    "agentMessages:",
    agentMessages.length,
  );
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const session = memoryStore.sessions.get(sessionId);
    if (!session) {
      console.error("[DB] Session not found for save:", sessionId);
      return;
    }
    let name = session.name;
    if (name === "New Chat") {
      const derivedName = deriveSessionName(agentMessages);
      if (derivedName) name = derivedName;
    }
    memoryStore.sessions.set(sessionId, {
      ...session,
      agentMessages,
      name,
      updatedAt: Date.now(),
    });
    console.log("[DB] saveSession complete (memory)");
    return;
  }

  const session = await db.get("sessions", sessionId);
  if (!session) {
    console.error("[DB] Session not found for save:", sessionId);
    return;
  }
  let name = session.name;
  if (name === "New Chat") {
    const derivedName = deriveSessionName(agentMessages);
    if (derivedName) name = derivedName;
  }
  await db.put("sessions", {
    ...session,
    agentMessages,
    name,
    updatedAt: Date.now(),
  });
  console.log("[DB] saveSession complete");
}

export async function renameSession(
  ns: StorageNamespace,
  sessionId: string,
  name: string,
): Promise<void> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const session = memoryStore.sessions.get(sessionId);
    if (session) {
      memoryStore.sessions.set(sessionId, { ...session, name });
    }
    return;
  }

  const session = await db.get("sessions", sessionId);
  if (session) {
    await db.put("sessions", { ...session, name });
  }
}

export async function deleteSession(
  ns: StorageNamespace,
  sessionId: string,
): Promise<void> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    memoryStore.sessions.delete(sessionId);
    return;
  }

  await db.delete("sessions", sessionId);
}

export async function getOrCreateCurrentSession(
  ns: StorageNamespace,
  workbookId: string,
): Promise<ChatSession> {
  const sessions = await listSessions(ns, workbookId);
  if (sessions.length > 0) {
    const session = sessions[0];
    if (!session.agentMessages) session.agentMessages = [];
    return session;
  }
  return createSession(ns, workbookId);
}

export async function saveVfsFiles(
  ns: StorageNamespace,
  sessionId: string,
  files: { path: string; data: Uint8Array }[],
): Promise<void> {
  console.log("[DB] saveVfsFiles:", sessionId, "files:", files.length);
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    // Delete existing files for this session
    for (const [key, file] of memoryStore.vfsFiles) {
      if (file.sessionId === sessionId) {
        memoryStore.vfsFiles.delete(key);
      }
    }
    // Add new files
    for (const f of files) {
      const id = `${sessionId}:${f.path}`;
      memoryStore.vfsFiles.set(id, {
        id,
        sessionId,
        path: f.path,
        data: f.data,
      });
    }
    return;
  }

  const tx = db.transaction("vfsFiles", "readwrite");
  const store = tx.store;
  const existing = await store.index("sessionId").getAllKeys(sessionId);
  for (const key of existing) {
    await store.delete(key);
  }
  for (const f of files) {
    await store.add({
      id: `${sessionId}:${f.path}`,
      sessionId,
      path: f.path,
      data: f.data,
    });
  }
  await tx.done;
}

export async function loadVfsFiles(
  ns: StorageNamespace,
  sessionId: string,
): Promise<{ path: string; data: Uint8Array }[]> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const rows = [...memoryStore.vfsFiles.values()]
      .filter(f => f.sessionId === sessionId);
    console.log("[DB] loadVfsFiles (memory):", sessionId, "files:", rows.length);
    return rows.map((r) => ({ path: r.path, data: r.data }));
  }

  const rows = await db.getAllFromIndex("vfsFiles", "sessionId", sessionId);
  console.log("[DB] loadVfsFiles:", sessionId, "files:", rows.length);
  return rows.map((r) => ({ path: r.path, data: r.data }));
}

export async function deleteVfsFiles(
  ns: StorageNamespace,
  sessionId: string,
): Promise<void> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    for (const [key, file] of memoryStore.vfsFiles) {
      if (file.sessionId === sessionId) {
        memoryStore.vfsFiles.delete(key);
      }
    }
    return;
  }

  const tx = db.transaction("vfsFiles", "readwrite");
  const keys = await tx.store.index("sessionId").getAllKeys(sessionId);
  for (const key of keys) {
    await tx.store.delete(key);
  }
  await tx.done;
}

export async function saveSkillFiles(
  ns: StorageNamespace,
  skillName: string,
  files: { path: string; data: Uint8Array }[],
): Promise<void> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    for (const [key, file] of memoryStore.skillFiles) {
      if (file.skillName === skillName) {
        memoryStore.skillFiles.delete(key);
      }
    }
    for (const f of files) {
      const id = `${skillName}:${f.path}`;
      memoryStore.skillFiles.set(id, {
        id,
        skillName,
        path: f.path,
        data: f.data,
      });
    }
    return;
  }

  const tx = db.transaction("skillFiles", "readwrite");
  const store = tx.store;
  const existing = await store.index("skillName").getAllKeys(skillName);
  for (const key of existing) {
    await store.delete(key);
  }
  for (const f of files) {
    await store.add({
      id: `${skillName}:${f.path}`,
      skillName,
      path: f.path,
      data: f.data,
    });
  }
  await tx.done;
}

export async function loadSkillFiles(
  ns: StorageNamespace,
  skillName: string,
): Promise<{ path: string; data: Uint8Array }[]> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const rows = [...memoryStore.skillFiles.values()]
      .filter(f => f.skillName === skillName);
    return rows.map((r) => ({ path: r.path, data: r.data }));
  }

  const rows = await db.getAllFromIndex("skillFiles", "skillName", skillName);
  return rows.map((r) => ({ path: r.path, data: r.data }));
}

export async function loadAllSkillFiles(
  ns: StorageNamespace,
): Promise<{ skillName: string; path: string; data: Uint8Array }[]> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const rows = [...memoryStore.skillFiles.values()];
    return rows.map((r) => ({
      skillName: r.skillName,
      path: r.path,
      data: r.data,
    }));
  }

  const rows = await db.getAll("skillFiles");
  return rows.map((r) => ({
    skillName: r.skillName,
    path: r.path,
    data: r.data,
  }));
}

export async function deleteSkillFiles(
  ns: StorageNamespace,
  skillName: string,
): Promise<void> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    for (const [key, file] of memoryStore.skillFiles) {
      if (file.skillName === skillName) {
        memoryStore.skillFiles.delete(key);
      }
    }
    return;
  }

  const tx = db.transaction("skillFiles", "readwrite");
  const keys = await tx.store.index("skillName").getAllKeys(skillName);
  for (const key of keys) {
    await tx.store.delete(key);
  }
  await tx.done;
}

export async function listSkillNames(ns: StorageNamespace): Promise<string[]> {
  const db = await getDb(ns);

  // NOTE: Memory fallback
  if (!db) {
    const names = new Set([...memoryStore.skillFiles.values()].map((r) => r.skillName));
    return [...names].sort();
  }

  const rows = await db.getAll("skillFiles");
  const names = new Set(rows.map((r) => r.skillName));
  return [...names].sort();
}
