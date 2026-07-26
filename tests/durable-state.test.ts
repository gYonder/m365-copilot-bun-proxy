import { describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DurableStateStore } from "../src/proxy/durable-state";
import { ResponseStore } from "../src/proxy/response-store";
import { SubstrateSessionStore } from "../src/proxy/substrate-session-store";
import type { WrapperOptions } from "../src/proxy/types";

describe("durable continuation metadata", () => {
  test("survives process-store reconstruction without storing content", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "m365-state-"));
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
});
