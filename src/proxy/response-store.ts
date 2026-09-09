import type {
  JsonObject,
  StoredOpenAiResponseRecord,
  WrapperOptions,
} from "./types";
import { createHash } from "node:crypto";
import { cloneJsonValue, nowUnix } from "./utils";
import { DurableStateStore } from "./durable-state";
import type { BridgeObservability } from "./observability";
import { ToolLedger } from "./tool-ledger";

type ConversationLinkEntry = {
  conversationId: string;
  expiresAtUtc: number;
};

type RequestHashEntry = {
  expiresAtUtc: number;
  conversationId: string | null;
  response: JsonObject;
};

export type StoredReplayResult = {
  conversationId: string | null;
  response: JsonObject | null;
};

type TaskDeadlineEntry = {
  deadlineMs: number;
  expiresAtUtc: number;
};

type ToolLedgerEntry = {
  ledger: ToolLedger;
  expiresAtUtc: number;
};

type ProtocolFailureCycleEntry = {
  phase: "retry_available" | "retry_claimed" | "exhausted";
  conversationId: string | null;
  response: JsonObject;
  expiresAtUtc: number;
};

export type ProtocolFailureCycleClaim =
  | { kind: "primary" }
  | {
      kind: "regenerate";
      conversationId: string | null;
      response: JsonObject;
    }
  | {
      kind: "replay";
      conversationId: string | null;
      response: JsonObject;
    };

export type ToolLedgerTaskScopeResolution =
  | { kind: "none" }
  | { kind: "resolved"; taskId: string }
  | { kind: "conflict" };

const RequestHashGuardTtlMs = 60_000;
const ProtocolFailureCycleTtlMs = 60_000;
const MaxStoredResponses = 1_024;
const MaxRequestHashes = 2_048;

export class ResponseStore {
  private readonly entries = new Map<string, StoredOpenAiResponseRecord>();
  private readonly conversationLinks = new Map<string, ConversationLinkEntry>();
  private readonly requestHashes = new Map<string, RequestHashEntry>();
  private readonly protocolReplayBodies = new Map<string, JsonObject>();
  private readonly taskDeadlines = new Map<string, TaskDeadlineEntry>();
  private readonly toolLedgers = new Map<string, ToolLedgerEntry>();
  private readonly inFlightResponses = new Map<string, Promise<Response>>();
  private readonly protocolFailureCycles = new Map<
    string,
    ProtocolFailureCycleEntry
  >();

  constructor(
    private readonly options: WrapperOptions,
    private readonly durable = new DurableStateStore(),
    private readonly observability: BridgeObservability | null = null,
    private readonly nowMs: () => number = Date.now,
  ) {
    const now = this.nowMs();
    for (const [id, entry] of Object.entries(this.durable.state.responses)) {
      if (entry.expiresAtUtc > now) this.conversationLinks.set(id, entry);
    }
    for (const [key, entry] of Object.entries(this.durable.state.toolLedgers)) {
      if (entry.expiresAtUtc <= now) continue;
      try {
        const recovered = ToolLedger.recover(entry.serialized, now, {
          now: () => Date.now(),
          pendingTtlMs: this.toolLedgerPendingTtlMs(),
        });
        this.recordToolLedgerRecovery(recovered.outcomes);
        if (recovered.ledger.size > 0) {
          this.toolLedgers.set(key, {
            ledger: recovered.ledger,
            expiresAtUtc: entry.expiresAtUtc,
          });
        }
      } catch {
        delete this.durable.state.toolLedgers[key];
      }
    }
    this.durable.save();
  }

  set(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
    taskDeadlineMs: number | null = null,
    contextInputTokens: number | null = null,
    contextWindowId: string | null = null,
  ): void {
    this.storeResponse(
      responseId,
      response,
      conversationId,
      taskDeadlineMs,
      contextInputTokens,
      contextWindowId,
      true,
    );
  }

  setTransient(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
    taskDeadlineMs: number | null = null,
  ): void {
    this.storeResponse(
      responseId,
      response,
      conversationId,
      taskDeadlineMs,
      null,
      null,
      false,
    );
  }

  private storeResponse(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
    taskDeadlineMs: number | null,
    contextInputTokens: number | null,
    contextWindowId: string | null,
    persistConversationLink: boolean,
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
    this.trimOldest(this.entries, MaxStoredResponses, "response");

    if (conversationId?.trim()) {
      this.conversationLinks.set(responseId, {
        conversationId: conversationId.trim(),
        expiresAtUtc: record.expiresAtUtc,
      });
      this.trimOldest(this.conversationLinks, MaxStoredResponses, "response_link");
      if (persistConversationLink) {
        this.durable.state.responses[responseId] = {
          conversationId: conversationId.trim(),
          expiresAtUtc: record.expiresAtUtc,
          contextInputTokens: record.contextInputTokens ?? undefined,
          contextWindowId: record.contextWindowId,
        };
        this.durable.save();
      }
    }
  }

  tryGetContextUsage(responseId: string): {
    inputTokens: number;
    windowId: string | null;
  } | null {
    this.purgeExpired();
    const entry = this.entries.get(responseId);
    if (entry?.contextInputTokens !== null && entry?.contextInputTokens !== undefined) {
      return {
        inputTokens: entry.contextInputTokens,
        windowId: entry.contextWindowId,
      };
    }
    const durableEntry = this.durable.state.responses[responseId];
    if (
      !durableEntry ||
      durableEntry.expiresAtUtc <= Date.now() ||
      durableEntry.contextInputTokens === undefined
    ) {
      return null;
    }
    return {
      inputTokens: durableEntry.contextInputTokens,
      windowId: durableEntry.contextWindowId ?? null,
    };
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
    const now = this.nowMs();
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
    this.trimOldest(this.taskDeadlines, MaxStoredResponses, "task_deadline");
    return deadlineMs;
  }

  getOrCreateToolLedger(taskKey: string): ToolLedger {
    this.purgeExpired();
    const key = durableKey(taskKey);
    const existing = this.toolLedgers.get(key);
    if (existing) {
      this.recordToolLedgerRecovery(existing.ledger.recoverExpired(Date.now()));
      if (existing.ledger.size === 0) {
        this.toolLedgers.delete(key);
        delete this.durable.state.toolLedgers[key];
      } else {
        this.toolLedgers.delete(key);
        this.toolLedgers.set(key, existing);
        return existing.ledger;
      }
    }

    const ledger = new ToolLedger({
      pendingTtlMs: this.toolLedgerPendingTtlMs(),
    });
    this.toolLedgers.set(key, {
      ledger,
      expiresAtUtc: this.resolveExpiryMs(),
    });
    this.trimToolLedgers();
    return ledger;
  }

  saveToolLedger(taskKey: string, ledger: ToolLedger): void {
    const key = durableKey(taskKey);
    if (ledger.size === 0) {
      this.toolLedgers.delete(key);
      if (delete this.durable.state.toolLedgers[key]) {
        this.durable.save();
      }
      return;
    }
    const expiresAtUtc = this.resolveExpiryMs();
    this.toolLedgers.set(key, { ledger, expiresAtUtc });
    this.durable.state.toolLedgers[key] = {
      serialized: ledger.serialize(),
      expiresAtUtc,
    };
    this.trimToolLedgers();
    this.durable.save();
  }

  resolveToolLedgerTaskScope(
    callIds: Iterable<string>,
    previousResponseId: string | null,
    fallbackTaskId: string,
  ): ToolLedgerTaskScopeResolution {
    const ids = Array.from(new Set(callIds))
      .map((id) => id.trim())
      .filter(Boolean);
    if (ids.length === 0) {
      return { kind: "none" };
    }
    this.purgeExpired();
    const taskIdsForResponse = new Set<string>();
    if (previousResponseId?.trim()) {
      for (const entry of this.toolLedgers.values()) {
        this.recordToolLedgerRecovery(entry.ledger.recoverExpired(Date.now()));
        const taskId = entry.ledger.getTaskIdForResponse(previousResponseId);
        if (taskId) {
          taskIdsForResponse.add(taskId);
        }
      }
    }
    if (taskIdsForResponse.size > 1) {
      return { kind: "conflict" };
    }
    const taskId =
      taskIdsForResponse.values().next().value ?? fallbackTaskId.trim();
    if (!taskId) {
      return { kind: "none" };
    }
    const entry = this.toolLedgers.get(durableKey(taskId));
    if (!entry) {
      return { kind: "none" };
    }
    this.recordToolLedgerRecovery(entry.ledger.recoverExpired(Date.now()));
    const matchedTaskIds = ids.map((callId) =>
      entry.ledger.getTaskIdForCall(callId),
    );
    if (matchedTaskIds.every((matchedTaskId) => matchedTaskId === null)) {
      return taskIdsForResponse.size === 1
        ? { kind: "conflict" }
        : { kind: "none" };
    }
    return matchedTaskIds.every((matchedTaskId) => matchedTaskId === taskId)
      ? { kind: "resolved", taskId }
      : { kind: "conflict" };
  }

  tryDelete(responseId: string): boolean {
    this.purgeExpired();
    const deletedEntry = this.entries.delete(responseId);
    const deletedLink = this.conversationLinks.delete(responseId);
    const deletedDurable = delete this.durable.state.responses[responseId];
    if (deletedDurable) {
      this.durable.save();
    }
    return deletedEntry || deletedLink || deletedDurable;
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
    const entry = {
      conversationId: conversationId.trim(),
      expiresAtUtc: this.resolveExpiryMs(),
    };
    this.conversationLinks.set(responseId, entry);
    this.durable.state.responses[responseId] = {
      ...this.durable.state.responses[responseId],
      ...entry,
    };
    this.durable.save();
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

  tryGetProtocolReplay(identityKey: string): StoredReplayResult | null {
    const key = identityKey.trim();
    if (!key) return null;
    this.purgeExpired();
    const entry = this.durable.state.replays[key];
    if (!entry) return null;
    if (entry.expiresAtUtc <= Date.now()) {
      this.protocolReplayBodies.delete(key);
      delete this.durable.state.replays[key];
      this.durable.save();
      return null;
    }
    const response = this.protocolReplayBodies.get(key);
    if (!response) {
      return {
        conversationId: entry.conversationId,
        response: null,
      };
    }
    this.protocolReplayBodies.delete(key);
    this.protocolReplayBodies.set(key, response);
    return {
      conversationId: entry.conversationId,
      response: cloneJsonValue(response),
    };
  }

  forgetProtocolReplay(identityKey: string): void {
    const key = identityKey.trim();
    if (!key) return;
    this.protocolReplayBodies.delete(key);
    if (this.durable.state.replays[key]) {
      delete this.durable.state.replays[key];
      this.durable.save();
    }
  }

  rememberCompletedProtocolTurn(
    identityKey: string,
    conversationId: string | null,
    response: JsonObject,
  ): void {
    const key = identityKey.trim();
    if (!key) return;
    this.purgeExpired();
    const expiresAtUtc = this.resolveExpiryMs();
    this.protocolReplayBodies.set(key, cloneJsonValue(response));
    this.durable.state.replays[key] = {
      conversationId: conversationId?.trim() || null,
      expiresAtUtc,
      responseFingerprint: createHash("sha256")
        .update(JSON.stringify(response))
        .digest("hex"),
    };
    while (this.protocolReplayBodies.size > MaxStoredResponses) {
      const oldest = this.protocolReplayBodies.keys().next();
      if (oldest.done) break;
      this.protocolReplayBodies.delete(oldest.value);
      delete this.durable.state.replays[oldest.value];
      this.recordEviction("protocol_replay", "lru");
    }
    this.durable.save();
  }

  claimProtocolFailureCycle(identityKey: string): ProtocolFailureCycleClaim {
    const key = identityKey.trim();
    if (!key) return { kind: "primary" };
    this.purgeExpired();
    const entry = this.protocolFailureCycles.get(key);
    if (!entry) return { kind: "primary" };
    if (entry.phase === "retry_available") {
      this.protocolFailureCycles.delete(key);
      this.protocolFailureCycles.set(key, {
        ...entry,
        phase: "retry_claimed",
        expiresAtUtc: this.nowMs() + ProtocolFailureCycleTtlMs,
      });
      return {
        kind: "regenerate",
        conversationId: entry.conversationId,
        response: cloneJsonValue(entry.response),
      };
    }
    return {
      kind: "replay",
      conversationId: entry.conversationId,
      response: cloneJsonValue(entry.response),
    };
  }

  rememberRetryableProtocolFailure(
    identityKey: string,
    conversationId: string | null,
    response: JsonObject,
  ): void {
    this.rememberProtocolFailure(
      identityKey,
      conversationId,
      response,
      "retry_available",
    );
  }

  rememberExhaustedProtocolFailure(
    identityKey: string,
    conversationId: string | null,
    response: JsonObject,
  ): void {
    this.rememberProtocolFailure(
      identityKey,
      conversationId,
      response,
      "exhausted",
    );
  }

  clearProtocolFailureCycle(identityKey: string): void {
    const key = identityKey.trim();
    if (key) this.protocolFailureCycles.delete(key);
  }

  tryGetInFlightResponse(identityKey: string): Promise<Response> | null {
    return this.inFlightResponses.get(identityKey.trim()) ?? null;
  }

  registerInFlightResponse(
    identityKey: string,
    promise: Promise<Response>,
  ): Promise<Response> {
    const key = identityKey.trim();
    if (!key) return promise;
    const existing = this.inFlightResponses.get(key);
    if (existing) return existing;
    this.inFlightResponses.set(key, promise);
    const cleanup = () => {
      if (this.inFlightResponses.get(key) === promise) {
        this.inFlightResponses.delete(key);
      }
    };
    void promise.then(cleanup, cleanup);
    return promise;
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
    this.trimOldest(this.requestHashes, MaxRequestHashes, "legacy_replay");
  }

  private resolveExpiryMs(): number {
    const ttlMinutes = this.options.conversationTtlMinutes;
    if (ttlMinutes <= 0) {
      return Number.MAX_SAFE_INTEGER;
    }
    return Date.now() + ttlMinutes * 60_000;
  }

  private purgeExpired(): void {
    let durableReplayChanged = false;
    const now = this.nowMs();

    if (this.entries.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.entries.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.entries.delete(id);
          this.conversationLinks.delete(id);
          this.recordEviction("response", "ttl");
        }
      }
    }

    if (this.conversationLinks.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.conversationLinks.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.conversationLinks.delete(id);
          this.recordEviction("response_link", "ttl");
        }
      }
    }

    if (this.requestHashes.size > 0) {
      const now = Date.now();
      for (const [hash, entry] of this.requestHashes.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.requestHashes.delete(hash);
          this.recordEviction("legacy_replay", "ttl");
        }
      }
    }

    if (this.taskDeadlines.size > 0) {
      for (const [key, entry] of this.taskDeadlines.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.taskDeadlines.delete(key);
          this.recordEviction("task_deadline", "ttl");
        }
      }
    }

    if (this.toolLedgers.size > 0) {
      for (const [key, entry] of this.toolLedgers.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.toolLedgers.delete(key);
          this.recordEviction("tool_ledger", "ttl");
        }
      }
    }

    for (const [key, entry] of this.protocolFailureCycles.entries()) {
      if (
        entry.expiresAtUtc <= now &&
        !(
          entry.phase === "retry_claimed" &&
          this.inFlightResponses.has(key)
        )
      ) {
        this.protocolFailureCycles.delete(key);
        this.recordEviction("protocol_failure_cycle", "ttl");
      }
    }

    for (const [key] of this.protocolReplayBodies.entries()) {
      const entry = this.durable.state.replays[key];
      if (!entry || entry.expiresAtUtc <= now) {
        this.protocolReplayBodies.delete(key);
      }
    }
    for (const [key, entry] of Object.entries(this.durable.state.replays)) {
      if (entry.expiresAtUtc <= now) {
        delete this.durable.state.replays[key];
        this.protocolReplayBodies.delete(key);
        durableReplayChanged = true;
        this.recordEviction("protocol_replay", "ttl");
      }
    }
    for (const [key, entry] of Object.entries(this.durable.state.toolLedgers)) {
      if (entry.expiresAtUtc <= now) {
        delete this.durable.state.toolLedgers[key];
        durableReplayChanged = true;
        this.recordEviction("tool_ledger", "ttl");
      }
    }
    if (durableReplayChanged) {
      this.durable.save();
    }
  }

  private trimOldest<K, V>(
    entries: Map<K, V>,
    maxEntries: number,
    store: string,
  ): void {
    while (entries.size > maxEntries) {
      const oldest = entries.keys().next();
      if (oldest.done) return;
      entries.delete(oldest.value);
      this.recordEviction(store, "lru");
    }
  }

  private trimToolLedgers(): void {
    while (this.toolLedgers.size > MaxStoredResponses) {
      const oldest = this.toolLedgers.keys().next();
      if (oldest.done) return;
      this.toolLedgers.delete(oldest.value);
      delete this.durable.state.toolLedgers[oldest.value];
      this.recordEviction("tool_ledger", "lru");
    }
  }

  private rememberProtocolFailure(
    identityKey: string,
    conversationId: string | null,
    response: JsonObject,
    phase: ProtocolFailureCycleEntry["phase"],
  ): void {
    const key = identityKey.trim();
    if (!key) return;
    this.purgeExpired();
    this.protocolFailureCycles.delete(key);
    this.protocolFailureCycles.set(key, {
      phase,
      conversationId: conversationId?.trim() || null,
      response: cloneJsonValue(response),
      expiresAtUtc: this.nowMs() + ProtocolFailureCycleTtlMs,
    });
    this.trimOldest(
      this.protocolFailureCycles,
      MaxRequestHashes,
      "protocol_failure_cycle",
    );
  }

  private recordToolLedgerRecovery(
    outcomes: Array<{ failure: { code: string; reason: string }; released: true }>,
  ): void {
    for (const outcome of outcomes) {
      this.observability?.record("tool_ledger_recovery", {
        source: "tool_ledger",
        code: outcome.failure.code,
        reason: outcome.failure.reason,
        released: outcome.released,
      });
    }
  }

  private toolLedgerPendingTtlMs(): number {
    return Math.max(1, this.options.conversationTtlMinutes) * 60_000;
  }

  private recordEviction(store: string, reason: "ttl" | "lru"): void {
    this.observability?.record("store_eviction", { store, reason });
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

function durableKey(key: string): string {
  return createHash("sha256").update(key).digest("hex");
}
