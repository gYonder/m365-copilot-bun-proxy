import type { WrapperOptions } from "./types";
import { createHash } from "node:crypto";
import { DurableStateStore } from "./durable-state";

const MaxConversationEntries = 1_024;

export class ConversationStore {
  private readonly entries = new Map<
    string,
    { conversationId: string; expiresAtUtc: number }
  >();
  constructor(
    private readonly options: WrapperOptions,
    private readonly durable = new DurableStateStore(),
  ) {
    const now = Date.now();
    for (const [key, entry] of Object.entries(this.durable.state.conversations)) {
      if (entry.expiresAtUtc > now) this.entries.set(key, entry);
    }
  }

  tryGet(key: string): string | null {
    this.purgeExpired();
    const hashedKey = durableKey(key);
    const entry = this.entries.get(key) ?? this.entries.get(hashedKey);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.entries.delete(key);
      this.entries.delete(hashedKey);
      return null;
    }
    this.entries.delete(key);
    this.entries.delete(hashedKey);
    this.entries.set(key, entry);
    return entry.conversationId;
  }

  set(key: string, conversationId: string): void {
    if (!key.trim() || !conversationId.trim()) {
      return;
    }
    const ttlMinutes = this.options.conversationTtlMinutes;
    const expiresAtUtc =
      ttlMinutes <= 0
        ? Number.MAX_SAFE_INTEGER
        : Date.now() + ttlMinutes * 60_000;
    this.entries.delete(key);
    this.entries.set(key, { conversationId, expiresAtUtc });
    while (this.entries.size > MaxConversationEntries) {
      const oldest = this.entries.keys().next();
      if (oldest.done) {
        break;
      }
      this.entries.delete(oldest.value);
    }
    this.durable.state.conversations[durableKey(key)] = {
      conversationId,
      expiresAtUtc,
    };
    this.durable.save();
  }

  private purgeExpired(): void {
    if (this.entries.size === 0) {
      return;
    }
    const now = Date.now();
    for (const [key, value] of this.entries.entries()) {
      if (value.expiresAtUtc <= now) {
        this.entries.delete(key);
      }
    }
  }
}

function durableKey(key: string): string {
  return createHash("sha256").update(key).digest("hex");
}
