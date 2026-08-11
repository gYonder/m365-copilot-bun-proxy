import type { DurableStateStore } from "./durable-state";

const BaseCooldownMs = 60_000;
const MaxCooldownMs = 15 * 60_000;

export type RateCircuitStatus = {
  open: boolean;
  openUntilUtc: number;
  consecutive: number;
  lastStatusCode: number | null;
};

/** Persistent upstream 429 guard shared by all OpenAI-compatible routes. */
export class RateCircuitBreaker {
  constructor(private readonly durable: DurableStateStore) {}

  status(now = Date.now()): RateCircuitStatus {
    const state = this.durable.state.rateLimit;
    return {
      open: state.openUntilUtc > now,
      openUntilUtc: state.openUntilUtc,
      consecutive: state.consecutive,
      lastStatusCode: state.lastStatusCode,
    };
  }

  retryAfterSeconds(now = Date.now()): number {
    return Math.max(1, Math.ceil((this.durable.state.rateLimit.openUntilUtc - now) / 1000));
  }

  allow(now = Date.now()): boolean {
    return !this.status(now).open;
  }

  record(statusCode: number, now = Date.now()): void {
    const state = this.durable.state.rateLimit;
    if (statusCode === 429) {
      state.consecutive += 1;
      state.lastStatusCode = 429;
      const cooldown = Math.min(
        MaxCooldownMs,
        BaseCooldownMs * 2 ** Math.max(0, state.consecutive - 1),
      );
      state.openUntilUtc = Math.max(state.openUntilUtc, now + cooldown);
      this.durable.save();
      return;
    }
    if (statusCode >= 200 && statusCode < 300 && state.consecutive > 0) {
      state.consecutive = 0;
      state.openUntilUtc = 0;
      state.lastStatusCode = null;
      this.durable.save();
    }
  }
}
