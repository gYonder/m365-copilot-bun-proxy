import { describe, expect, test } from "bun:test";
import {
  SubstrateSessionStore,
  TurnQueueWaitError,
} from "../src/proxy/substrate-session-store";

describe("SubstrateSessionStore", () => {
  test("reuses session id for repeated turns in the same conversation", () => {
    const store = new SubstrateSessionStore(180);
    let created = 0;
    const createSessionId = () => `session-${++created}`;

    const first = store.getOrCreate("conversation-1", createSessionId, 1_000);
    const second = store.getOrCreate("conversation-1", createSessionId, 2_000);

    expect(first).toBe("session-1");
    expect(second).toBe("session-1");
    expect(created).toBe(1);
  });

  test("creates a new session id after ttl expiration", () => {
    const store = new SubstrateSessionStore(1);
    let created = 0;
    const createSessionId = () => `session-${++created}`;

    const first = store.getOrCreate("conversation-1", createSessionId, 0);
    const second = store.getOrCreate(
      "conversation-1",
      createSessionId,
      61_000,
    );

    expect(first).toBe("session-1");
    expect(second).toBe("session-2");
  });

  test("can bind a known session id to a conversation id", () => {
    const store = new SubstrateSessionStore(180);
    store.set("conversation-2", "session-bound", 1_000);

    const resolved = store.getOrCreate("conversation-2", () => "session-new", 2_000);
    expect(resolved).toBe("session-bound");
  });

  test("serializes turns for the same conversation", async () => {
    const store = new SubstrateSessionStore(180);
    const events: string[] = [];
    let releaseFirst!: () => void;
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });

    const first = store.runExclusive("conversation-1", async () => {
      events.push("first:start");
      await firstGate;
      events.push("first:end");
    });
    const second = store.runExclusive("conversation-1", async () => {
      events.push("second:start");
      events.push("second:end");
    });

    await Bun.sleep(0);
    expect(events).toEqual(["first:start"]);
    releaseFirst();
    await Promise.all([first, second]);
    expect(events).toEqual([
      "first:start",
      "first:end",
      "second:start",
      "second:end",
    ]);
  });

  test("bounds inactive session entries", () => {
    const store = new SubstrateSessionStore(180, 2);
    store.set("conversation-1", "session-1", 1_000);
    store.set("conversation-2", "session-2", 2_000);
    store.set("conversation-3", "session-3", 3_000);

    expect(
      store.getOrCreate("conversation-1", () => "session-recreated", 4_000),
    ).toBe("session-recreated");
    expect(store.getOrCreate("conversation-3", () => "unexpected", 4_000)).toBe(
      "session-3",
    );
  });

  test("removes cancelled waiters without blocking later turns", async () => {
    const store = new SubstrateSessionStore(180);
    let releaseFirst!: () => void;
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    const first = store.runExclusive("conversation-1", () => firstGate);
    const controller = new AbortController();
    const cancelled = store.runExclusive(
      "conversation-1",
      async () => "unexpected",
      controller.signal,
    );
    controller.abort();

    await expect(cancelled).rejects.toEqual(
      new TurnQueueWaitError("cancelled"),
    );
    releaseFirst();
    await first;
    expect(
      await store.runExclusive("conversation-1", async () => "next"),
    ).toBe("next");
  });

  test("trims entries again after active turns finish", async () => {
    const store = new SubstrateSessionStore(180, 1);
    let releaseFirst!: () => void;
    let releaseSecond!: () => void;
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    const secondGate = new Promise<void>((resolve) => {
      releaseSecond = resolve;
    });
    const first = store.runExclusive("conversation-1", async () => {
      store.set("conversation-1", "session-1", 1_000);
      await firstGate;
    });
    const second = store.runExclusive("conversation-2", async () => {
      store.set("conversation-2", "session-2", 2_000);
      await secondGate;
    });

    await Bun.sleep(0);
    releaseFirst();
    await first;
    releaseSecond();
    await second;

    expect(store.getOrCreate("conversation-2", () => "unexpected", 3_000)).toBe(
      "session-2",
    );
    expect(
      store.getOrCreate("conversation-1", () => "session-recreated", 3_000),
    ).toBe("session-recreated");
  });

  test("uses the same lock when a conversation id is rebound to a session", async () => {
    const store = new SubstrateSessionStore(180);
    const sessionId = store.getOrCreate("conversation-old", () => "session-1");
    store.set("conversation-new", sessionId);
    const reboundSessionId = store.getOrCreate("conversation-new");
    const events: string[] = [];
    let releaseFirst!: () => void;
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });

    const first = store.runExclusive(sessionId, async () => {
      events.push("first");
      await firstGate;
    });
    const second = store.runExclusive(reboundSessionId, async () => {
      events.push("second");
    });

    await Bun.sleep(0);
    expect(events).toEqual(["first"]);
    releaseFirst();
    await Promise.all([first, second]);
    expect(events).toEqual(["first", "second"]);
  });

  test("rejects excess per-conversation waiters", async () => {
    const store = new SubstrateSessionStore(180);
    let releaseFirst!: () => void;
    const first = store.runExclusive(
      "session-1",
      () =>
        new Promise<void>((resolve) => {
          releaseFirst = resolve;
        }),
    );
    const controllers = Array.from({ length: 32 }, () => new AbortController());
    const queued = controllers.map((controller) =>
      store.runExclusive("session-1", async () => {}, controller.signal),
    );

    await expect(
      store.runExclusive("session-1", async () => {}),
    ).rejects.toEqual(new TurnQueueWaitError("overloaded"));

    controllers.forEach((controller) => controller.abort());
    await Promise.allSettled(queued);
    releaseFirst();
    await first;
  });
});
