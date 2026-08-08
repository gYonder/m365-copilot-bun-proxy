import { describe, expect, test } from "bun:test";
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DurableStateStore } from "../src/proxy/durable-state";
import { ResponseStore } from "../src/proxy/response-store";
import { SubstrateSessionStore } from "../src/proxy/substrate-session-store";
import type { WrapperOptions } from "../src/proxy/types";

describe("durable continuation metadata", () => {
  test("survives process-store reconstruction without storing content", () => {
    const dir = createTempDir("m365-state");
    const file = path.join(dir, "continuation.json");
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const first = new DurableStateStore(file);
    const responses = new ResponseStore(options, first);
    const conversations = new ConversationStore(options, first);
    const sessions = new SubstrateSessionStore(180, 512, first);

    responses.set(
      "resp_1",
      { id: "resp_1", output_text: "must not persist" },
      "conv_1",
      null,
      123,
      "window_1",
    );
    responses.setConversationLink("resp_1", "conv_1");
    conversations.set("private-user-key", "conv_1");
    sessions.getOrCreate("conv_1", () => "sess_1");

    const second = new DurableStateStore(file);
    const resumedResponses = new ResponseStore(options, second);
    const resumedConversations = new ConversationStore(options, second);
    const resumedSessions = new SubstrateSessionStore(180, 512, second);

    expect(resumedResponses.tryGetConversationLink("resp_1")).toBe("conv_1");
    expect(resumedResponses.tryGetContextUsage("resp_1")).toEqual({
      inputTokens: 123,
      windowId: "window_1",
    });
    expect(resumedConversations.tryGet("private-user-key")).toBe("conv_1");
    expect(resumedSessions.getOrCreate("conv_1")).toBe("sess_1");

    const persisted = readFileSync(file, "utf8");
    expect(persisted).not.toContain("must not persist");
    expect(persisted).not.toContain("private-user-key");
    expect(statSync(file).mode & 0o777).toBe(0o600);
    rmSync(dir, { recursive: true, force: true });
  });

  test("persists explicit completed protocol replays across reconstruction", () => {
    const dir = createTempDir("m365-replay");
    const file = path.join(dir, "continuation.json");
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const first = new ResponseStore(options, new DurableStateStore(file));
    const completed = {
      id: "resp_protocol_1",
      object: "response",
      status: "completed",
      output: [],
      output_text: "completed protocol result",
    };

    first.rememberCompletedProtocolTurn(
      "protocol:durable-turn",
      "conv_protocol_1",
      completed,
    );

    expect(first.tryGetProtocolReplay("protocol:durable-turn")).toEqual({
      conversationId: "conv_protocol_1",
      response: completed,
    });

    const resumed = new ResponseStore(options, new DurableStateStore(file));
    expect(resumed.tryGetProtocolReplay("protocol:durable-turn")).toEqual({
      conversationId: "conv_protocol_1",
      response: null,
    });
    expect(statSync(file).mode & 0o777).toBe(0o600);
    rmSync(dir, { recursive: true, force: true });
  });

  test("persists bounded tool ledger metadata without tool results", () => {
    const dir = createTempDir("m365-ledger");
    const file = path.join(dir, "continuation.json");
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const first = new ResponseStore(options, new DurableStateStore(file));
    const ledger = first.getOrCreateToolLedger("task-key");
    const issued = ledger.issueCalls({
      taskId: "task-key",
      responseId: "response-1",
      requestProfileKey: "profile-1",
      calls: [
        {
          call_id: "call-1",
          name: "read_file",
          type: "function",
          arguments: { path: "README.md" },
        },
      ],
      round: 1,
    });
    expect(issued.ok).toBeTrue();
    first.saveToolLedger("task-key", ledger);

    const resumed = new ResponseStore(options, new DurableStateStore(file));
    expect(resumed.getOrCreateToolLedger("task-key").get("call-1")).not.toBeNull();
    expect(readFileSync(file, "utf8")).not.toContain("tool result body");
    rmSync(dir, { recursive: true, force: true });
  });

  test("removes expired durable protocol replays", () => {
    const dir = createTempDir("m365-replay-expired");
    const file = path.join(dir, "continuation.json");
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const durable = new DurableStateStore(file);
    durable.state.replays["protocol:expired-turn"] = {
      conversationId: "conv_expired",
      expiresAtUtc: Date.now() - 1,
      responseFingerprint: "expired",
    };
    durable.save();

    const resumed = new ResponseStore(options, new DurableStateStore(file));
    expect(resumed.tryGetProtocolReplay("protocol:expired-turn")).toBeNull();
    expect(readFileSync(file, "utf8")).not.toContain("protocol:expired-turn");
    rmSync(dir, { recursive: true, force: true });
  });

  test("quarantines the old body-bearing state format", () => {
    const dir = createTempDir("m365-old-state");
    const file = path.join(dir, "continuation.json");
    writeFileSync(
      file,
      JSON.stringify({
        version: 1,
        replays: {
          "protocol:old": {
            conversationId: "conv_old",
            response: { output_text: "old body" },
            expiresAtUtc: Date.now() + 60_000,
          },
        },
      }),
    );

    const store = new DurableStateStore(file);

    expect(store.state.replays).toEqual({});
    expect(existsSync(file)).toBeFalse();
    expect(quarantineFiles(file)).toHaveLength(1);
    rmSync(dir, { recursive: true, force: true });
  });

  test("quarantines an unparseable state file", () => {
    const dir = createTempDir("m365-corrupt-state");
    const file = path.join(dir, "continuation.json");
    writeFileSync(file, "{not-json");

    const store = new DurableStateStore(file);

    expect(store.state.replays).toEqual({});
    expect(existsSync(file)).toBeFalse();
    expect(quarantineFiles(file)).toHaveLength(1);
    rmSync(dir, { recursive: true, force: true });
  });

  test("starts clean without quarantining an absent state file", () => {
    const dir = createTempDir("m365-absent-state");
    const file = path.join(dir, "continuation.json");

    const store = new DurableStateStore(file);

    expect(store.state.replays).toEqual({});
    expect(existsSync(file)).toBeFalse();
    expect(quarantineFiles(file)).toHaveLength(0);
    rmSync(dir, { recursive: true, force: true });
  });

  test("never overwrites an existing quarantine file", () => {
    const dir = createTempDir("m365-quarantine-collision");
    const file = path.join(dir, "continuation.json");
    const fixedNow = 1_700_000_000_000;
    const existingQuarantine = `${file}.incompatible-${fixedNow}`;
    writeFileSync(file, JSON.stringify({ version: 1 }));
    writeFileSync(existingQuarantine, "keep this evidence");

    const originalDateNow = Date.now;
    Date.now = () => fixedNow;
    try {
      new DurableStateStore(file);
    } finally {
      Date.now = originalDateNow;
    }

    expect(readFileSync(existingQuarantine, "utf8")).toBe("keep this evidence");
    expect(existsSync(`${existingQuarantine}-1`)).toBeTrue();
    rmSync(dir, { recursive: true, force: true });
  });

  test("does not persist prompts, output, or inline image data in replay state", () => {
    const dir = createTempDir("m365-replay-content");
    const file = path.join(dir, "continuation.json");
    const store = new ResponseStore(
      { conversationTtlMinutes: 180 } as WrapperOptions,
      new DurableStateStore(file),
    );
    store.rememberCompletedProtocolTurn("protocol:content", "conv_content", {
      id: "resp_content",
      input: [{ role: "user", content: "secret prompt text" }],
      output: [{ type: "message", text: "secret model output" }],
      image: "data:image/png;base64,secret-image-bytes",
    });

    const persisted = readFileSync(file, "utf8");
    expect(persisted).not.toContain("secret prompt text");
    expect(persisted).not.toContain("secret model output");
    expect(persisted).not.toContain("data:image/");
    expect(statSync(file).mode & 0o777).toBe(0o600);
    rmSync(dir, { recursive: true, force: true });
  });

  test("round-trips the new state envelope and durable content", () => {
    const dir = createTempDir("m365-round-trip");
    const file = path.join(dir, "continuation.json");
    const first = new DurableStateStore(file);
    first.state.responses.response_1 = {
      conversationId: "conv_1",
      expiresAtUtc: Date.now() + 60_000,
      contextInputTokens: 42,
      contextWindowId: "window_1",
    };
    first.state.replays.replay_1 = {
      conversationId: "conv_1",
      expiresAtUtc: Date.now() + 60_000,
      responseFingerprint: "fingerprint",
    };
    first.save();

    const second = new DurableStateStore(file);

    expect(second.state).toEqual(first.state);
    expect(second.state.schema_version).toBe(1);
    expect(second.state.contract_version).toBe(2);
    expect(second.state.adapter_route).toBe("substrate-coding");
    expect(statSync(file).mode & 0o777).toBe(0o600);
    rmSync(dir, { recursive: true, force: true });
  });

  test("reports an unreplayable durable completion without a response body", () => {
    const dir = createTempDir("m365-unreplayable");
    const file = path.join(dir, "continuation.json");
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const first = new ResponseStore(options, new DurableStateStore(file));
    first.rememberCompletedProtocolTurn(
      "protocol:unreplayable",
      "conv_unreplayable",
      { id: "resp_unreplayable", output_text: "not persisted" },
    );

    const resumed = new ResponseStore(options, new DurableStateStore(file));
    const replay = resumed.tryGetProtocolReplay("protocol:unreplayable");

    expect(replay).toEqual({
      conversationId: "conv_unreplayable",
      response: null,
    });
    rmSync(dir, { recursive: true, force: true });
  });
});

function createTempDir(prefix: string): string {
  return mkdtempSync(path.join(process.cwd(), `.${prefix}-`));
}

function quarantineFiles(file: string): string[] {
  const dir = path.dirname(file);
  const base = path.basename(file);
  return readdirSync(dir).filter((name) =>
    name.startsWith(`${base}.incompatible-`),
  );
}
