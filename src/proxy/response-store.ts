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
};

type ReplayBodyEntry = {
  expiresAtUtc: number;
  response: JsonObject;
};

const RequestHashGuardTtlMs = 60_000;

export class ResponseStore {
  private readonly entries = new Map<string, StoredOpenAiResponseRecord>();
  private readonly conversationLinks = new Map<string, ConversationLinkEntry>();
  private readonly requestHashes = new Map<string, RequestHashEntry>();
  private readonly replayBodies = new Map<string, ReplayBodyEntry>();

  constructor(private readonly options: WrapperOptions) {}

  set(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
  ): void {
    if (!responseId.trim()) {
      return;
    }
    this.purgeExpired();
    const record: StoredOpenAiResponseRecord = {
      responseId,
      createdAtUnix: readCreatedAt(response),
      response: cloneJsonValue(response),
      conversationId: conversationId?.trim() ? conversationId : null,
      expiresAtUtc: this.resolveExpiryMs(),
    };
    this.entries.set(responseId, record);

    if (conversationId?.trim()) {
      this.conversationLinks.set(responseId, {
        conversationId: conversationId.trim(),
        expiresAtUtc: record.expiresAtUtc,
      });
      // Cache the completed body for idempotent replay of a byte-identical
      // retry inside the guard window (see rememberReplayResponse).
      this.rememberReplayResponse(conversationId.trim(), response);
    }
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

  hasRecentRequestHash(requestHash: string): boolean {
    const normalizedHash = requestHash.trim();
    if (!normalizedHash) {
      return false;
    }
    this.purgeExpired();
    const entry = this.requestHashes.get(normalizedHash);
    if (!entry) {
      return false;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.requestHashes.delete(normalizedHash);
      return false;
    }
    return true;
  }

  getRecentRequestHashConversationId(requestHash: string): string | null {
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
    return entry.conversationId;
  }

  rememberRequestHash(
    requestHash: string,
    conversationId: string | null = null,
  ): void {
    const normalizedHash = requestHash.trim();
    if (!normalizedHash) {
      return;
    }
    this.purgeExpired();
    this.requestHashes.set(normalizedHash, {
      expiresAtUtc: Date.now() + RequestHashGuardTtlMs,
      conversationId: conversationId?.trim() ? conversationId.trim() : null,
    });
  }

  // Cache the real completed response body per conversation so a byte-identical
  // retry inside the guard window can be answered with the same body (crucially,
  // any function_call output items) instead of an empty suppressed message.
  // Keyed by conversationId — the one join key available both here (via set())
  // and at the guard hit site (via getRecentRequestHashConversationId). For a
  // byte-identical retry the client never received turn N's response (that is
  // why it retries), so it cannot have advanced the conversation past N, and the
  // latest cached body for the conversation is turn N's. Guard TTL; kept out of
  // the entries map so it never surfaces in list().
  rememberReplayResponse(conversationId: string, response: JsonObject): void {
    const normalizedId = conversationId.trim();
    if (!normalizedId) {
      return;
    }
    // Only cache responses that carry useful output (a function_call item or
    // non-empty assistant text). Empty / tool-less bodies are left uncached so a
    // byte-identical retry falls through to a fresh upstream attempt instead of
    // being permanently answered with an empty replay — this preserves the
    // stochastic "retry until the model emits the forced tool call" flow while
    // still replaying real tool/text turns idempotently.
    if (!responseHasUsefulOutput(response)) {
      return;
    }
    this.purgeExpired();
    this.replayBodies.set(normalizedId, {
      expiresAtUtc: Date.now() + RequestHashGuardTtlMs,
      response: cloneJsonValue(response),
    });
  }

  tryGetReplayResponse(conversationId: string): JsonObject | null {
    const normalizedId = conversationId.trim();
    if (!normalizedId) {
      return null;
    }
    this.purgeExpired();
    const entry = this.replayBodies.get(normalizedId);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.replayBodies.delete(normalizedId);
      return null;
    }
    return cloneJsonValue(entry.response);
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

    if (this.replayBodies.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.replayBodies.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.replayBodies.delete(id);
        }
      }
    }
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

// A response is worth caching for idempotent replay only if it carries output a
// retry would want back: a function_call (tool) item, or a message with
// non-empty text. Empty / tool-less bodies return false so they are not cached.
function responseHasUsefulOutput(response: JsonObject): boolean {
  const output = response.output;
  if (!Array.isArray(output)) {
    return false;
  }
  for (const item of output) {
    if (typeof item !== "object" || item === null || Array.isArray(item)) {
      continue;
    }
    const typed = item as Record<string, unknown>;
    const itemType = typeof typed.type === "string" ? typed.type : "";
    if (itemType === "function_call") {
      return true;
    }
    if (itemType === "message" || itemType === "") {
      const content = typed.content;
      if (!Array.isArray(content)) {
        continue;
      }
      for (const part of content) {
        if (typeof part !== "object" || part === null || Array.isArray(part)) {
          continue;
        }
        const partTyped = part as Record<string, unknown>;
        if (
          typeof partTyped.text === "string" &&
          partTyped.text.trim().length > 0
        ) {
          return true;
        }
      }
    }
  }
  return false;
}
