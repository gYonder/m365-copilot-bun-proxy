import { TurnQueueWaitError } from "./substrate-session-store";

type LimiterWaiter = {
  resolve: () => void;
  reject: (error: TurnQueueWaitError) => void;
  signal: AbortSignal | null;
  abortHandler: (() => void) | null;
  timer: ReturnType<typeof setTimeout> | null;
};

/**
 * Caps how many Substrate turns may execute concurrently across every
 * conversation. The per-conversation lock in {@link SubstrateSessionStore}
 * serializes turns within a single conversation; this limiter bounds the total
 * number of in-flight upstream WebSocket turns so a burst of parallel
 * conversations cannot overwhelm Substrate (which responds to overload with
 * opaque disconnects). A limit of 0 disables the cap entirely.
 */
export class ConcurrencyLimiter {
  private available: number;
  private readonly waiters: LimiterWaiter[] = [];

  constructor(
    private readonly limit: number,
    private readonly acquireTimeoutMs = 0,
  ) {
    this.available = limit > 0 ? limit : 0;
  }

  get enabled(): boolean {
    return this.limit > 0;
  }

  /**
   * Acquires a slot, returning a release function. When disabled the release is
   * a no-op. Throws {@link TurnQueueWaitError} on cancellation, task deadline,
   * or acquire-timeout so callers can map it to 499/504/429 respectively.
   */
  async acquire(
    signal: AbortSignal | null = null,
    deadlineMs: number | null = null,
  ): Promise<() => void> {
    if (this.limit <= 0) {
      return () => {};
    }
    if (signal?.aborted) {
      throw new TurnQueueWaitError("cancelled");
    }
    if (deadlineMs !== null && deadlineMs <= Date.now()) {
      throw new TurnQueueWaitError("deadline");
    }
    if (this.available > 0) {
      this.available -= 1;
      return this.makeRelease();
    }

    return new Promise<() => void>((resolve, reject) => {
      const waiter: LimiterWaiter = {
        resolve: () => {
          this.cleanupWaiter(waiter);
          resolve(this.makeRelease());
        },
        reject,
        signal,
        abortHandler: null,
        timer: null,
      };
      const rejectAndRemove = (reason: "cancelled" | "deadline" | "overloaded") => {
        const index = this.waiters.indexOf(waiter);
        if (index >= 0) {
          this.waiters.splice(index, 1);
        }
        this.cleanupWaiter(waiter);
        reject(new TurnQueueWaitError(reason));
      };
      if (signal) {
        waiter.abortHandler = () => rejectAndRemove("cancelled");
        signal.addEventListener("abort", waiter.abortHandler, { once: true });
      }
      const deadlineWaitMs =
        deadlineMs !== null ? Math.max(0, deadlineMs - Date.now()) : null;
      const timeoutWaitMs = this.acquireTimeoutMs > 0 ? this.acquireTimeoutMs : null;
      const effectiveWaitMs =
        deadlineWaitMs !== null && timeoutWaitMs !== null
          ? Math.min(deadlineWaitMs, timeoutWaitMs)
          : (deadlineWaitMs ?? timeoutWaitMs);
      if (effectiveWaitMs !== null) {
        const reason: "deadline" | "overloaded" =
          timeoutWaitMs !== null &&
          (deadlineWaitMs === null || timeoutWaitMs <= deadlineWaitMs)
            ? "overloaded"
            : "deadline";
        waiter.timer = setTimeout(() => rejectAndRemove(reason), effectiveWaitMs);
      }
      this.waiters.push(waiter);
    });
  }

  private makeRelease(): () => void {
    let released = false;
    return () => {
      if (released) {
        return;
      }
      released = true;
      const next = this.waiters.shift();
      if (next) {
        next.resolve();
        return;
      }
      this.available += 1;
    };
  }

  private cleanupWaiter(waiter: LimiterWaiter): void {
    if (waiter.abortHandler && waiter.signal) {
      waiter.signal.removeEventListener("abort", waiter.abortHandler);
    }
    if (waiter.timer) {
      clearTimeout(waiter.timer);
    }
  }
}
