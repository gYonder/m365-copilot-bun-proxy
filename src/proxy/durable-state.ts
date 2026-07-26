import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";

type DurableResponseEntry = {
  conversationId: string;
  expiresAtUtc: number;
  contextInputTokens?: number;
  contextWindowId?: string | null;
};

type DurableConversationEntry = {
  conversationId: string;
  expiresAtUtc: number;
};

type DurableSessionEntry = {
  sessionId: string;
  expiresAtUtc: number;
};

export type DurableState = {
  version: 1;
  responses: Record<string, DurableResponseEntry>;
  conversations: Record<string, DurableConversationEntry>;
  sessions: Record<string, DurableSessionEntry>;
};

const emptyState = (): DurableState => ({
  version: 1,
  responses: {},
  conversations: {},
  sessions: {},
});

export function durableStatePath(): string | null {
  const configured = process.env.M365_RUNTIME_STATE_PATH?.trim();
  return configured || null;
}

export class DurableStateStore {
  readonly path: string | null;
  state: DurableState;
  constructor(filePath = durableStatePath()) {
    this.path = filePath;
    this.state = this.load();
  }

  private load(): DurableState {
    if (!this.path || !existsSync(this.path)) return emptyState();
    try {
      return parseState(JSON.parse(readFileSync(this.path, "utf8")));
    } catch {
      return emptyState();
    }
  }

  save(): void {
    if (!this.path) return;
    const now = Date.now();
    for (const collection of [this.state.responses, this.state.conversations, this.state.sessions]) {
      for (const [key, value] of Object.entries(collection)) {
        if (value.expiresAtUtc <= now) delete collection[key];
      }
      const keys = Object.keys(collection);
      for (const key of keys.slice(0, Math.max(0, keys.length - 1024))) delete collection[key];
    }
    const dir = path.dirname(this.path);
    mkdirSync(dir, { recursive: true, mode: 0o700 });
    const temp = `${this.path}.${process.pid}.tmp`;
    writeFileSync(temp, JSON.stringify(this.state) + "\n", { mode: 0o600 });
    chmodSync(temp, 0o600);
    renameSync(temp, this.path);
    chmodSync(this.path, 0o600);
  }
}

function parseState(value: unknown): DurableState {
  if (!isRecord(value) || value.version !== 1) {
    return emptyState();
  }
  return {
    version: 1,
    responses: parseResponseEntries(value.responses),
    conversations: parseConversationEntries(value.conversations),
    sessions: parseSessionEntries(value.sessions),
  };
}

function parseResponseEntries(value: unknown): Record<string, DurableResponseEntry> {
  const output: Record<string, DurableResponseEntry> = {};
  if (!isRecord(value)) return output;
  for (const [key, entry] of Object.entries(value)) {
    if (!isRecord(entry)) continue;
    const conversationId = readString(entry.conversationId);
    const expiresAtUtc = readPositiveNumber(entry.expiresAtUtc);
    if (!conversationId || expiresAtUtc === null) continue;
    const contextInputTokens = readPositiveNumber(entry.contextInputTokens);
    const contextWindowId =
      entry.contextWindowId === null ? null : readString(entry.contextWindowId);
    output[key] = {
      conversationId,
      expiresAtUtc,
      ...(contextInputTokens === null ? {} : { contextInputTokens }),
      ...(contextWindowId === undefined ? {} : { contextWindowId }),
    };
  }
  return output;
}

function parseConversationEntries(
  value: unknown,
): Record<string, DurableConversationEntry> {
  const output: Record<string, DurableConversationEntry> = {};
  if (!isRecord(value)) return output;
  for (const [key, entry] of Object.entries(value)) {
    if (!isRecord(entry)) continue;
    const conversationId = readString(entry.conversationId);
    const expiresAtUtc = readPositiveNumber(entry.expiresAtUtc);
    if (conversationId && expiresAtUtc !== null) {
      output[key] = { conversationId, expiresAtUtc };
    }
  }
  return output;
}

function parseSessionEntries(value: unknown): Record<string, DurableSessionEntry> {
  const output: Record<string, DurableSessionEntry> = {};
  if (!isRecord(value)) return output;
  for (const [key, entry] of Object.entries(value)) {
    if (!isRecord(entry)) continue;
    const sessionId = readString(entry.sessionId);
    const expiresAtUtc = readPositiveNumber(entry.expiresAtUtc);
    if (sessionId && expiresAtUtc !== null) {
      output[key] = { sessionId, expiresAtUtc };
    }
  }
  return output;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function readPositiveNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? value
    : null;
}
