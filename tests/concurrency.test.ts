import { describe, expect, test } from "bun:test";
import { ConcurrencyLimiter } from "../src/proxy/concurrency";
import { TurnQueueWaitError } from "../src/proxy/substrate-session-store";

describe("ConcurrencyLimiter", () => {
  test("is disabled when the limit is zero", async () => {
    const limiter = new ConcurrencyLimiter(0);
    expect(limiter.enabled).toBeFalse();
    const releaseA = await limiter.acquire();
    const releaseB = await limiter.acquire();
    // No cap: both resolve immediately and release is a safe no-op.
    releaseA();
    releaseB();
  });

  test("caps concurrent holders and hands the slot to the next waiter", async () => {
    const limiter = new ConcurrencyLimiter(1);
    const releaseFirst = await limiter.acquire();

    let secondAcquired = false;
    const secondPromise = limiter.acquire().then((release) => {
      secondAcquired = true;
      return release;
    });

    await Promise.resolve();
    expect(secondAcquired).toBeFalse();

    releaseFirst();
    const releaseSecond = await secondPromise;
    expect(secondAcquired).toBeTrue();
    releaseSecond();
  });

  test("reports queue depth and hands slots to waiters in FIFO order", async () => {
    const limiter = new ConcurrencyLimiter(1);
    const releaseFirst = await limiter.acquire();
    const order: string[] = [];

    const second = limiter.acquire().then((release) => {
      order.push("second");
      return release;
    });
    const third = limiter.acquire().then((release) => {
      order.push("third");
      return release;
    });

    expect(limiter.activeCount).toBe(1);
    expect(limiter.queueDepth).toBe(2);

    releaseFirst();
    const releaseSecond = await second;
    expect(order).toEqual(["second"]);
    expect(limiter.queueDepth).toBe(1);

    releaseSecond();
    const releaseThird = await third;
    expect(order).toEqual(["second", "third"]);
    expect(limiter.queueDepth).toBe(0);

    releaseThird();
    expect(limiter.activeCount).toBe(0);
  });

  test("rejects a queued waiter when its signal aborts", async () => {
    const limiter = new ConcurrencyLimiter(1);
    const release = await limiter.acquire();
    const controller = new AbortController();
    const waiting = limiter.acquire(controller.signal);
    controller.abort();

    await expect(waiting).rejects.toBeInstanceOf(TurnQueueWaitError);
    release();
  });

  test("rejects immediately when the deadline has already passed", async () => {
    const limiter = new ConcurrencyLimiter(1);
    const release = await limiter.acquire();
    await expect(limiter.acquire(null, Date.now() - 1)).rejects.toBeInstanceOf(
      TurnQueueWaitError,
    );
    release();
  });

  test("times out a queued waiter after the acquire timeout as overloaded", async () => {
    const limiter = new ConcurrencyLimiter(1, 20);
    const release = await limiter.acquire();
    let caught: unknown = null;
    try {
      await limiter.acquire();
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(TurnQueueWaitError);
    expect((caught as TurnQueueWaitError).reason).toBe("overloaded");
    expect(limiter.queueDepth).toBe(0);
    release();
  });
});
