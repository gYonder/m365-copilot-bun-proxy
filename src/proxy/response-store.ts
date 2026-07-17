import type {
  JsonObject,
  StoredOpenAiResponseRecord,
  WrapperOptions,
} from "./types";
import { cloneJsonValue, nowUnix } from "./utils";

type ConversationLinkEntry = {
  conversationId: string;
  expiresAtUtc: number;
};

type RequestHashEntry = {
  expiresAtUtc: number;
  conversationId: string | null;
  response: JsonObject;
};

type TaskDeadlineEntry = {
  deadlineMs: number;
  expiresAtUtc: number;
};

const RequestHashGuardTtlMs = 60_000;
const MaxStoredResponses = 1_024;
const MaxRequestHashes = 2_048;

export class ResponseStore {
  private readonly entries = new Map<string, StoredOpenAiResponseRecord>();
  private readonly conversationLinks = new Map<string, ConversationLinkEntry>();
  private readonly requestHashes = new Map<string, RequestHashEntry>();
  private readonly taskDeadlines = new Map<string, TaskDeadlineEntry>();

  constructor(private readonly options: WrapperOptions) {}

  set(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
    taskDeadlineMs: number | null = null,
    contextInputTokens: number | null = null,
    contextWindowId: string | null = null,
  ): void {
    if (!responseId.trim()) {
      return;
    }
    this.purgeExpired();
    const inferredContext = readContextUsage(response);
    const record: StoredOpenAiResponseRecord = {
      responseId,
      createdAtUnix: readCreatedAt(response),
      response: cloneJsonValue(response),
      conversationId: conversationId?.trim() ? conversationId : null,
      expiresAtUtc: this.resolveExpiryMs(),
      taskDeadlineMs,
      contextInputTokens: contextInputTokens ?? inferredContext.inputTokens,
      contextWindowId: contextWindowId ?? inferredContext.windowId,
    };
    this.entries.set(responseId, record);
    trimOldest(this.entries, MaxStoredResponses);

    if (conversationId?.trim()) {
      this.conversationLinks.set(responseId, {
        conversationId: conversationId.trim(),
        expiresAtUtc: record.expiresAtUtc,
      });
      trimOldest(this.conversationLinks, MaxStoredResponses);
    }
  }

  tryGetContextUsage(responseId: string): {
    inputTokens: number;
    windowId: string | null;
  } | null {
    this.purgeExpired();
    const entry = this.entries.get(responseId);
    if (!entry || entry.contextInputTokens === null) return null;
    return { inputTokens: entry.contextInputTokens, windowId: entry.contextWindowId };
  }

  tryGet(responseId: string): JsonObject | null {
    this.purgeExpired();
    const entry = this.entries.get(responseId);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.entries.delete(responseId);
      this.conversationLinks.delete(responseId);
      return null;
    }
    return cloneJsonValue(entry.response);
  }

  tryGetTaskDeadline(responseId: string): number | null {
    this.purgeExpired();
    const entry = this.entries.get(responseId);
    if (!entry || entry.expiresAtUtc <= Date.now()) {
      return null;
    }
    return entry.taskDeadlineMs;
  }

  getOrCreateTaskDeadline(
    taskKey: string,
    createDeadline: () => number,
  ): number {
    const normalizedKey = taskKey.trim();
    if (!normalizedKey) {
      return createDeadline();
    }
    this.purgeExpired();
    const existing = this.taskDeadlines.get(normalizedKey);
    const now = Date.now();
    if (existing && existing.deadlineMs > now && existing.expiresAtUtc > now) {
      this.taskDeadlines.delete(normalizedKey);
      this.taskDeadlines.set(normalizedKey, existing);
      return existing.deadlineMs;
    }
    const deadlineMs = createDeadline();
    this.taskDeadlines.set(normalizedKey, {
      deadlineMs,
      expiresAtUtc: Math.max(deadlineMs, Date.now()) + RequestHashGuardTtlMs,
    });
    trimOldest(this.taskDeadlines, MaxStoredResponses);
    return deadlineMs;
  }

  tryDelete(responseId: string): boolean {
    this.purgeExpired();
    const deletedEntry = this.entries.delete(responseId);
    const deletedLink = this.conversationLinks.delete(responseId);
    return deletedEntry || deletedLink;
  }

  list(limit: number): {
    data: JsonObject[];
    hasMore: boolean;
    firstId: string | null;
    lastId: string | null;
  } {
    this.purgeExpired();
    const normalizedLimit = normalizeLimit(limit);
    const sorted = [...this.entries.values()].sort(
      (a, b) => b.createdAtUnix - a.createdAtUnix,
    );
    const total = sorted.length;
    const selected = sorted.slice(0, normalizedLimit);
    return {
      data: selected.map((entry) => cloneJsonValue(entry.response)),
      hasMore: total > selected.length,
      firstId: selected.length > 0 ? selected[0].responseId : null,
      lastId:
        selected.length > 0 ? selected[selected.length - 1].responseId : null,
    };
  }

  setConversationLink(responseId: string, conversationId: string): void {
    if (!responseId.trim() || !conversationId.trim()) {
      return;
    }
    this.purgeExpired();
    this.conversationLinks.set(responseId, {
      conversationId: conversationId.trim(),
      expiresAtUtc: this.resolveExpiryMs(),
    });
  }

  tryGetConversationLink(responseId: string): string | null {
    this.purgeExpired();
    const entry = this.conversationLinks.get(responseId);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.conversationLinks.delete(responseId);
      return null;
    }
    return entry.conversationId;
  }

  tryGetRequestReplay(
    requestHash: string,
  ): { conversationId: string | null; response: JsonObject } | null {
    const normalizedHash = requestHash.trim();
    if (!normalizedHash) {
      return null;
    }
    this.purgeExpired();
    const entry = this.requestHashes.get(normalizedHash);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.requestHashes.delete(normalizedHash);
      return null;
    }
    this.requestHashes.delete(normalizedHash);
    this.requestHashes.set(normalizedHash, entry);
    return {
      conversationId: entry.conversationId,
      response: cloneJsonValue(entry.response),
    };
  }

  rememberCompletedRequest(
    requestHash: string,
    conversationId: string | null,
    response: JsonObject,
  ): void {
    const normalizedHash = requestHash.trim();
    if (!normalizedHash) {
      return;
    }
    this.purgeExpired();
    this.requestHashes.set(normalizedHash, {
      expiresAtUtc: Date.now() + RequestHashGuardTtlMs,
      conversationId: conversationId?.trim() ? conversationId.trim() : null,
      response: cloneJsonValue(response),
    });
    trimOldest(this.requestHashes, MaxRequestHashes);
  }

  private resolveExpiryMs(): number {
    const ttlMinutes = this.options.conversationTtlMinutes;
    if (ttlMinutes <= 0) {
      return Number.MAX_SAFE_INTEGER;
    }
    return Date.now() + ttlMinutes * 60_000;
  }

  private purgeExpired(): void {
    if (this.entries.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.entries.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.entries.delete(id);
          this.conversationLinks.delete(id);
        }
      }
    }

    if (this.conversationLinks.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.conversationLinks.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.conversationLinks.delete(id);
        }
      }
    }

    if (this.requestHashes.size > 0) {
      const now = Date.now();
      for (const [hash, entry] of this.requestHashes.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.requestHashes.delete(hash);
        }
      }
    }

    if (this.taskDeadlines.size > 0) {
      const now = Date.now();
      for (const [key, entry] of this.taskDeadlines.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.taskDeadlines.delete(key);
        }
      }
    }

  }
}

function readContextUsage(response: JsonObject): {
  inputTokens: number | null;
  windowId: string | null;
} {
  const usage = response.usage;
  if (!usage || typeof usage !== "object" || Array.isArray(usage)) {
    return { inputTokens: null, windowId: null };
  }
  const inputTokens = usage.x_m365_context_input_tokens;
  const windowId = usage.x_m365_context_window_id;
  return {
    inputTokens:
      typeof inputTokens === "number" && Number.isFinite(inputTokens)
        ? Math.max(0, Math.trunc(inputTokens))
        : null,
    windowId:
      typeof windowId === "string" && windowId.trim() ? windowId.trim() : null,
  };
}

function trimOldest<K, V>(entries: Map<K, V>, maxEntries: number): void {
  while (entries.size > maxEntries) {
    const oldest = entries.keys().next();
    if (oldest.done) {
      return;
    }
    entries.delete(oldest.value);
  }
}

function normalizeLimit(rawLimit: number): number {
  if (!Number.isFinite(rawLimit)) {
    return 20;
  }
  const rounded = Math.trunc(rawLimit);
  if (rounded <= 0) {
    return 20;
  }
  return rounded > 100 ? 100 : rounded;
}

function readCreatedAt(response: JsonObject): number {
  const value = response.created_at;
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return nowUnix();
}
