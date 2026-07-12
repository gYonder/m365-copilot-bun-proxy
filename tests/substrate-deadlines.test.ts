import { describe, expect, test } from "bun:test";
import { resolveSubstrateDeadlines } from "../src/proxy/clients";
import type { SubstrateOptions } from "../src/proxy/types";

function makeSubstrate(
  overrides: Partial<
    Pick<
      SubstrateOptions,
      | "invocationTimeoutSeconds"
      | "handshakeTimeoutSeconds"
      | "turnTimeoutSeconds"
    >
  >,
): SubstrateOptions {
  return {
    invocationTimeoutSeconds: 120,
    ...overrides,
  } as SubstrateOptions;
}

describe("resolveSubstrateDeadlines", () => {
  test("defaults both handshake and turn to the invocation timeout when unset", () => {
    const now = 1_000_000;
    const deadlines = resolveSubstrateDeadlines(
      makeSubstrate({ invocationTimeoutSeconds: 90 }),
      now,
    );
    expect(deadlines.invocationTimeoutMs).toBe(90_000);
    expect(deadlines.handshakeTimeoutMs).toBe(90_000);
    // turn deadline is absolute: now + turn budget.
    expect(deadlines.turnDeadlineMs).toBe(now + 90_000);
  });

  test("resolves handshake and whole-turn deadlines independently of each other", () => {
    const now = 5_000;
    const deadlines = resolveSubstrateDeadlines(
      makeSubstrate({
        invocationTimeoutSeconds: 120,
        handshakeTimeoutSeconds: 10,
        turnTimeoutSeconds: 300,
      }),
      now,
    );
    // Handshake timeout is short and does NOT inherit the whole-turn budget.
    expect(deadlines.handshakeTimeoutMs).toBe(10_000);
    // Whole-turn deadline is absolute and independent of the handshake timeout.
    expect(deadlines.turnDeadlineMs).toBe(now + 300_000);
    // The per-invocation socket timeout stays separate from both.
    expect(deadlines.invocationTimeoutMs).toBe(120_000);
    // Prove they are genuinely separate values, not the same number reused.
    expect(deadlines.handshakeTimeoutMs).not.toBe(deadlines.invocationTimeoutMs);
    expect(deadlines.turnDeadlineMs - now).not.toBe(deadlines.handshakeTimeoutMs);
  });

  test("only one of handshake or turn overriding leaves the other at the invocation default", () => {
    const now = 0;
    const handshakeOnly = resolveSubstrateDeadlines(
      makeSubstrate({ invocationTimeoutSeconds: 120, handshakeTimeoutSeconds: 5 }),
      now,
    );
    expect(handshakeOnly.handshakeTimeoutMs).toBe(5_000);
    expect(handshakeOnly.turnDeadlineMs).toBe(now + 120_000);

    const turnOnly = resolveSubstrateDeadlines(
      makeSubstrate({ invocationTimeoutSeconds: 120, turnTimeoutSeconds: 600 }),
      now,
    );
    expect(turnOnly.handshakeTimeoutMs).toBe(120_000);
    expect(turnOnly.turnDeadlineMs).toBe(now + 600_000);
  });

  test("falls back to 120s invocation budget for a non-positive configured timeout", () => {
    const now = 42;
    const deadlines = resolveSubstrateDeadlines(
      makeSubstrate({ invocationTimeoutSeconds: 0 }),
      now,
    );
    expect(deadlines.invocationTimeoutMs).toBe(120_000);
    expect(deadlines.handshakeTimeoutMs).toBe(120_000);
    expect(deadlines.turnDeadlineMs).toBe(now + 120_000);
  });
});
