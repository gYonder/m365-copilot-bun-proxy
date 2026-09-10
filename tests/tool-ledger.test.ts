import { describe, expect, test } from "bun:test";
import {
  MAX_TOOL_CALLS_PER_RESPONSE,
  TOOL_LEDGER_PENDING_TTL_MS,
  ToolLedger,
  ToolLedgerCodecError,
  decideToolParallelism,
} from "../src/proxy/tool-ledger";

describe("ToolLedger", () => {
  test("records every issued call as pending", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    const result = issue(ledger, "call-1", { b: 2, a: 1 });

    expect(result.ok).toBe(true);
    expect(ledger.get("call-1")).toEqual({
      call_id: "call-1",
      name: "read_file",
      type: "function",
      canonical_arguments:
        "43258cff783fe7036d8a43033f830adfc60ec037382473548ac742b888292777",
      request_profile_key: "profile-1",
      issued_at: 1_000,
      status: "pending",
    });
  });

  test("accepts exactly one result and rejects the duplicate", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    issue(ledger, "call-1");

    expect(ledger.acceptResult("call-1", "profile-1", "small").ok).toBe(true);
    const duplicate = ledger.acceptResult("call-1", "profile-1", "large");
    expect(duplicate.ok).toBe(false);
    if (!duplicate.ok) {
      expect(duplicate.error.kind).toBe("duplicate_result");
      expect(duplicate.error.code).toBe("duplicate_suppressed");
    }
    expect(ledger.get("call-1")?.status).toBe("completed");
  });

  test("accepts reverse-order siblings but rejects superseded rounds", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    issue(ledger, "call-1");
    issue(ledger, "call-2", { second: true }, 1, "response-2");

    const unknown = ledger.acceptResult("missing", "profile-1");
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.error.kind).toBe("unknown_call_id");

    const crossProfile = ledger.acceptResult("call-1", "profile-2");
    expect(crossProfile.ok).toBe(false);
    if (!crossProfile.ok) expect(crossProfile.error.code).toBe("schema_invalid");

    expect(ledger.acceptResult("call-2", "profile-1").ok).toBe(true);
    expect(ledger.acceptResult("call-1", "profile-1").ok).toBe(true);

    const superseded = new ToolLedger({ now: () => 1_000 });
    issue(superseded, "old-call", { old: true }, 1, "response-old");
    issue(superseded, "new-call", { newer: true }, 2, "response-new");
    const outOfOrder = superseded.acceptResult("old-call", "profile-1");
    expect(outOfOrder.ok).toBe(false);
    if (!outOfOrder.ok) expect(outOfOrder.error.kind).toBe("out_of_order");
    expect(superseded.acceptResult("new-call", "profile-1").ok).toBe(true);
  });

  test("resets repetition detection after a different call", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    expect(issue(ledger, "call-1", { a: 1, b: 2 }).ok).toBe(true);
    expect(ledger.acceptResult("call-1", "profile-1").ok).toBe(true);
    expect(
      issue(
        ledger,
        "call-2",
        JSON.stringify({ b: 2, a: 1 }),
        2,
        "response-2",
      ).ok,
    ).toBe(true);
    expect(ledger.acceptResult("call-2", "profile-1").ok).toBe(true);

    expect(issue(ledger, "call-3", { path: "other.txt" }, 3, "response-3").ok).toBe(
      true,
    );
    expect(ledger.acceptResult("call-3", "profile-1").ok).toBe(true);
    expect(issue(ledger, "call-4", { a: 1, b: 2 }, 4, "response-4").ok).toBe(
      true,
    );
  });

  test("rejects a runaway sequence of consecutive identical calls", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    for (let index = 1; index <= 4; index += 1) {
      const callId = `call-${index}`;
      expect(
        issue(ledger, callId, { a: 1, b: 2 }, index, `response-${index}`).ok,
      ).toBe(true);
      expect(ledger.acceptResult(callId, "profile-1").ok).toBe(true);
    }

    const repeated = issue(ledger, "call-5", { a: 1, b: 2 }, 5, "response-5");
    expect(repeated.ok).toBe(false);
    if (!repeated.ok) {
      expect(repeated.error.kind).toBe("repetition_bound_exhausted");
      expect(repeated.error.code).toBe("fallback_exhausted");
    }
  });

  test("enforces per-response calls and per-task rounds", () => {
    const ledger = new ToolLedger({
      now: () => 1_000,
      maxCallsPerResponse: 2,
      maxRoundsPerTask: 2,
    });
    const first = ledger.issueCalls({
      taskId: "task-1",
      responseId: "response-1",
      requestProfileKey: "profile-1",
      calls: [call("call-1", { one: 1 }), call("call-2", { two: 2 })],
      round: 1,
    });
    expect(first.ok).toBe(true);
    const responseCap = issue(
      ledger,
      "call-3",
      { three: 3 },
      1,
      "response-1",
    );
    expect(responseCap.ok).toBe(false);
    if (!responseCap.ok) expect(responseCap.error.kind).toBe("call_cap_exceeded");

    expect(ledger.acceptResult("call-1", "profile-1").ok).toBe(true);
    expect(ledger.acceptResult("call-2", "profile-1").ok).toBe(true);
    expect(issue(ledger, "call-4", { four: 4 }, 2, "response-2").ok).toBe(
      true,
    );
    const roundCap = issue(
      ledger,
      "call-5",
      { five: 5 },
      3,
      "response-3",
    );
    expect(roundCap.ok).toBe(false);
    if (!roundCap.ok) expect(roundCap.error.kind).toBe("round_cap_exceeded");
  });

  test("does not evict pending calls when the entry cap is reached", () => {
    const ledger = new ToolLedger({
      now: () => 1_000,
      maxEntries: 2,
    });
    expect(issue(ledger, "call-1", { one: 1 }).ok).toBe(true);
    expect(issue(ledger, "call-2", { two: 2 }, 2, "response-2").ok).toBe(
      true,
    );

    const rejected = issue(ledger, "call-3", { three: 3 }, 3, "response-3");
    expect(rejected.ok).toBe(false);
    if (!rejected.ok) {
      expect(rejected.error.kind).toBe("ledger_capacity_exceeded");
    }
    expect(ledger.pendingCalls("task-1").map((entry) => entry.call_id)).toEqual([
      "call-1",
      "call-2",
    ]);
  });

  test("allows one corrective re-ask per malformed call", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    issue(ledger, "call-1");
    expect(ledger.requestCorrectiveReask("call-1", "profile-1").ok).toBe(true);
    const second = ledger.requestCorrectiveReask("call-1", "profile-1");
    expect(second.ok).toBe(false);
    if (!second.ok) {
      expect(second.error.kind).toBe("corrective_reask_exhausted");
      expect(second.error.code).toBe("fallback_exhausted");
    }
  });

  test("expires pending calls with an injected clock", () => {
    let now = 1_000;
    const ledger = new ToolLedger({ now: () => now });
    issue(ledger, "call-1");
    now += TOOL_LEDGER_PENDING_TTL_MS;

    const outcomes = ledger.recoverExpired();
    expect(outcomes).toHaveLength(1);
    expect(outcomes[0].failure.code).toBe("ambiguous_completion");
    expect(outcomes[0].released).toBe(true);
    expect(ledger.hasPending("task-1")).toBe(false);
  });

  test("recovers expired persisted calls without leaving a task pending", () => {
    const first = new ToolLedger({ now: () => 1_000 });
    issue(first, "call-1");
    const recovered = ToolLedger.recover(
      first.serialize(),
      1_000 + TOOL_LEDGER_PENDING_TTL_MS,
      { now: () => 1_000 + TOOL_LEDGER_PENDING_TTL_MS },
    );

    expect(recovered.outcomes[0].failure.code).toBe("ambiguous_completion");
    expect(recovered.ledger.hasPending("task-1")).toBe(false);
    expect(recovered.ledger.size).toBe(0);
  });

  test("round-trips the codec and rejects incompatible state", () => {
    const first = new ToolLedger({ now: () => 1_000 });
    issue(first, "call-1", { nested: { b: 2, a: 1 } });
    const serialized = first.serialize();
    const second = ToolLedger.deserialize(serialized, { now: () => 1_000 });
    expect(second.get("call-1")).toEqual(first.get("call-1"));

    const wrongVersion = serialized.replace(
      '"version":1',
      '"version":999',
    );
    expect(() => ToolLedger.deserialize(wrongVersion)).toThrow(
      ToolLedgerCodecError,
    );
    expect(() => ToolLedger.deserialize("{not-json")).toThrow(
      ToolLedgerCodecError,
    );
    expect(() =>
      ToolLedger.deserialize(
        JSON.stringify({ version: 1, entries: [], metadata: [] }),
      ),
    ).toThrow(ToolLedgerCodecError);
  });

  test("does not retain tool-result bodies", () => {
    const ledger = new ToolLedger({ now: () => 1_000 });
    issue(ledger, "call-1");
    const before = ledger.serialize().length;
    ledger.acceptResult("call-1", "profile-1", "x".repeat(1_000_000));
    const after = ledger.serialize();

    expect(after.length).toBeLessThan(before + 100);
    expect(after).not.toContain("x".repeat(1_000));
  });

  test("uses only structured metadata for parallelism", () => {
    expect(decideToolParallelism(undefined).reason).toBe(
      "independence_signal_unavailable",
    );
    expect(
      decideToolParallelism([
        {
          requestPermitsParallelCalls: true,
          independentlyReadOnly: true,
          consumesAnotherCallOutput: false,
        },
        {
          requestPermitsParallelCalls: true,
          independentlyReadOnly: true,
          consumesAnotherCallOutput: false,
        },
      ]).parallel,
    ).toBe(true);
    expect(
      decideToolParallelism([
        {
          requestPermitsParallelCalls: true,
          independentlyReadOnly: false,
          consumesAnotherCallOutput: false,
        },
      ]).reason,
    ).toBe("mutating_or_dependency_sensitive");
    expect(
      decideToolParallelism(
        Array.from({ length: MAX_TOOL_CALLS_PER_RESPONSE + 1 }, () => ({
          requestPermitsParallelCalls: true,
          independentlyReadOnly: true,
          consumesAnotherCallOutput: false,
        })),
      ).reason,
    ).toBe("call_bound_exceeded");
  });
});

function call(callId: string, argumentsValue: unknown) {
  return {
    call_id: callId,
    name: "read_file",
    type: "function",
    arguments: argumentsValue,
  };
}

function issue(
  ledger: ToolLedger,
  callId: string,
  argumentsValue: unknown = { path: "README.md" },
  round = 1,
  responseId = `response-${callId}`,
) {
  return ledger.issueCalls({
    taskId: "task-1",
    responseId,
    requestProfileKey: "profile-1",
    calls: [call(callId, argumentsValue)],
    round,
  });
}
