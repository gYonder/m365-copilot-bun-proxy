import { randomUUID } from "node:crypto";
import { DurableStateStore } from "./durable-state";

type SessionEntry = {
  sessionId: string;
  expiresAtUtc: number;
};

const DefaultMaxEntries = 512;
const MaxWaitersPerConversation = 32;

type TurnWaiter = {
  resolve: () => void;
  reject: (error: TurnQueueWaitError) => void;
  signal: AbortSignal | null;
  abortHandler: (() => void) | null;
  timer: ReturnType<typeof setTimeout> | null;
};

export class TurnQueueWaitError extends Error {
  constructor(readonly reason: "cancelled" | "deadline" | "overloaded") {
    super(
      reason === "cancelled"
        ? "Substrate turn cancelled while waiting for the conversation lock."
        : reason === "deadline"
          ? "M365 task-level deadline exceeded while waiting for the conversation lock."
          : "Too many Substrate turns are queued for this conversation.",
    );
  }
}

export class SubstrateSessionStore {
  private readonly entries = new Map<string, SessionEntry>();
  private readonly activeTurns = new Set<string>();
  private readonly turnQueues = new Map<string, TurnWaiter[]>();
  constructor(
    private readonly ttlMinutes: number,
    private readonly maxEntries: number = DefaultMaxEntries,
    private readonly durable = new DurableStateStore(),
  ) {
    const now = Date.now();
    for (const [key, entry] of Object.entries(this.durable.state.sessions)) {
      if (entry.expiresAtUtc > now) this.entries.set(key, entry);
    }
  }

  get activeTurnCount(): number {
    return this.activeTurns.size;
  }

  get queuedTurnCount(): number {
    let count = 0;
    for (const queue of this.turnQueues.values()) {
      count += queue.length;
    }
    return count;
  }

  getOrCreate(
    conversationId: string,
    createSessionId: () => string = () => randomUUID(),
    nowUtcMs: number = Date.now(),
  ): string {
    this.purgeExpired(nowUtcMs);
    const normalizedConversationId = conversationId.trim();
    if (!normalizedConversationId) {
      return createSessionId();
    }

    const existing = this.entries.get(normalizedConversationId);
    if (existing) {
      const refreshed = {
        sessionId: existing.sessionId,
        expiresAtUtc: this.computeExpiry(nowUtcMs),
      };
      this.entries.delete(normalizedConversationId);
      this.entries.set(normalizedConversationId, refreshed);
      if (
        this.shouldRefreshDurableSession(
          normalizedConversationId,
          existing.sessionId,
          nowUtcMs,
        )
      ) {
        this.durable.state.sessions[normalizedConversationId] = refreshed;
        this.durable.save();
      }
      return existing.sessionId;
    }

    const sessionId = createSessionId();
    const entry = {
      sessionId,
      expiresAtUtc: this.computeExpiry(nowUtcMs),
    };
    this.entries.set(normalizedConversationId, entry);
    this.durable.state.sessions[normalizedConversationId] = entry;
    this.durable.save();
    this.trimToLimit();
    return sessionId;
  }

  set(
    conversationId: string,
    sessionId: string,
    nowUtcMs: number = Date.now(),
  ): void {
    this.purgeExpired(nowUtcMs);
    const normalizedConversationId = conversationId.trim();
    const normalizedSessionId = sessionId.trim();
    if (!normalizedConversationId || !normalizedSessionId) {
      return;
    }
    const entry = {
      sessionId: normalizedSessionId,
      expiresAtUtc: this.computeExpiry(nowUtcMs),
    };
    this.entries.set(normalizedConversationId, entry);
    this.durable.state.sessions[normalizedConversationId] = entry;
    this.durable.save();
    this.trimToLimit();
  }

  async runExclusive<T>(
    conversationId: string,
    task: () => Promise<T>,
    signal: AbortSignal | null = null,
    deadlineMs: number | null = null,
  ): Promise<T> {
    const key = conversationId.trim();
    if (!key) {
      return task();
    }

    await this.acquireTurn(key, signal, deadlineMs);
    try {
      return await task();
    } finally {
      this.releaseTurn(key);
      this.trimToLimit();
    }
  }

  private computeExpiry(nowUtcMs: number): number {
    return this.ttlMinutes <= 0
      ? Number.MAX_SAFE_INTEGER
      : nowUtcMs + this.ttlMinutes * 60_000;
  }

  private shouldRefreshDurableSession(
    conversationId: string,
    sessionId: string,
    nowUtcMs: number,
  ): boolean {
    const persisted = this.durable.state.sessions[conversationId];
    if (!persisted || persisted.sessionId !== sessionId) {
      return true;
    }
    if (this.ttlMinutes <= 0) {
      return persisted.expiresAtUtc !== Number.MAX_SAFE_INTEGER;
    }
    const refreshWindowMs = (this.ttlMinutes * 60_000) / 2;
    return persisted.expiresAtUtc <= nowUtcMs + refreshWindowMs;
  }

  private purgeExpired(nowUtcMs: number): void {
    if (this.entries.size === 0) {
      return;
    }
    for (const [conversationId, entry] of this.entries.entries()) {
      if (entry.expiresAtUtc <= nowUtcMs) {
        this.entries.delete(conversationId);
      }
    }
  }

  private trimToLimit(): void {
    if (this.maxEntries <= 0) {
      return;
    }
    for (const conversationId of this.entries.keys()) {
      if (this.entries.size <= this.maxEntries) {
        return;
      }
      if (!this.activeTurns.has(conversationId)) {
        this.entries.delete(conversationId);
      }
    }
  }

  private acquireTurn(
    key: string,
    signal: AbortSignal | null,
    deadlineMs: number | null,
  ): Promise<void> {
    if (signal?.aborted) {
      return Promise.reject(new TurnQueueWaitError("cancelled"));
    }
    if (deadlineMs !== null && deadlineMs <= Date.now()) {
      return Promise.reject(new TurnQueueWaitError("deadline"));
    }
    if (!this.activeTurns.has(key)) {
      this.activeTurns.add(key);
      return Promise.resolve();
    }
    const existingQueue = this.turnQueues.get(key);
    if ((existingQueue?.length ?? 0) >= MaxWaitersPerConversation) {
      return Promise.reject(new TurnQueueWaitError("overloaded"));
    }

    return new Promise<void>((resolve, reject) => {
      const waiter: TurnWaiter = {
        resolve: () => {
          this.cleanupWaiter(waiter);
          resolve();
        },
        reject,
        signal,
        abortHandler: null,
        timer: null,
      };
      const rejectAndRemove = (reason: "cancelled" | "deadline") => {
        const queue = this.turnQueues.get(key);
        if (queue) {
          const index = queue.indexOf(waiter);
          if (index >= 0) {
            queue.splice(index, 1);
          }
          if (queue.length === 0) {
            this.turnQueues.delete(key);
          }
        }
        this.cleanupWaiter(waiter);
        reject(new TurnQueueWaitError(reason));
      };
      if (signal) {
        waiter.abortHandler = () => rejectAndRemove("cancelled");
        signal.addEventListener("abort", waiter.abortHandler, { once: true });
      }
      if (deadlineMs !== null) {
        waiter.timer = setTimeout(
          () => rejectAndRemove("deadline"),
          Math.max(0, deadlineMs - Date.now()),
        );
      }
      const queue = this.turnQueues.get(key) ?? [];
      queue.push(waiter);
      this.turnQueues.set(key, queue);
    });
  }

  private releaseTurn(key: string): void {
    const queue = this.turnQueues.get(key);
    const next = queue?.shift();
    if (queue && queue.length === 0) {
      this.turnQueues.delete(key);
    }
    if (next) {
      next.resolve();
      return;
    }
    this.activeTurns.delete(key);
  }

  private cleanupWaiter(waiter: TurnWaiter): void {
    if (waiter.abortHandler && waiter.signal) {
      waiter.signal.removeEventListener("abort", waiter.abortHandler);
    }
    if (waiter.timer) {
      clearTimeout(waiter.timer);
    }
  }
}
