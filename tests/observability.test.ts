import { describe, expect, test } from "bun:test";
import {
  BridgeObservability,
  sanitizeTelemetryObject,
} from "../src/proxy/observability";
import type { JsonObject } from "../src/proxy/types";
import { buildInvocationPayload } from "../src/proxy/clients";
import type { WrapperOptions } from "../src/proxy/types";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DurableStateStore } from "../src/proxy/durable-state";

describe("sanitized bridge observability", () => {
  test("redacts sensitive keys, bearer values, and authenticated URLs", () => {
    const signedUrl = "https://artifact.invalid/image?signature=secret";
    const sanitized = sanitizeTelemetryObject({
      outcome: "failed",
      authorization: "Bearer secret-token",
      prompt: "private prompt",
      accountId: "private-account",
      socket: "wss://example.invalid/chat?access_token=secret",
      nested: { cookie: "session=secret", artifactUrl: signedUrl },
      message: "failed with Bearer another-secret",
    });
    const serialized = JSON.stringify(sanitized);
    expect(sanitized.outcome).toBe("failed");
    expect(serialized).not.toContain("secret-token");
    expect(serialized).not.toContain("private prompt");
    expect(serialized).not.toContain("private-account");
    expect(serialized).not.toContain("access_token");
    expect(serialized).not.toContain(signedUrl);
    expect(serialized).not.toContain("another-secret");
  });

  test("keeps the first failed attempt visible after a successful retry", () => {
    const metrics = new BridgeObservability();
    metrics.record("attempt_outcome", { attempt: 1, success: false, statusCode: 502 });
    metrics.record("retry", { reason: "confab_give_up", retryCount: 1 });
    metrics.record("attempt_outcome", { attempt: 2, success: true, statusCode: 200 });
    const readiness = metrics.readiness();
    const summaries = readiness.events as JsonObject;
    expect((summaries.attempt_outcome as JsonObject).count).toBe(2);
    expect((summaries.retry as JsonObject).count).toBe(1);
    const recent = readiness.recentEvents as JsonObject[];
    expect(recent.map((event) => event.name)).toEqual([
      "attempt_outcome",
      "retry",
      "attempt_outcome",
    ]);
    expect((recent[0].fields as JsonObject).success).toBeFalse();
    expect((recent[2].fields as JsonObject).success).toBeTrue();
  });

  test("hashes supplied correlation material and never returns it verbatim", () => {
    const metrics = new BridgeObservability();
    const event = metrics.record("dedup_hit", { kind: "protocol_replay" }, "private-turn-id");
    expect(event.correlationId).toStartWith("corr_");
    expect(event.correlationId).not.toContain("private-turn-id");
    expect(event.correlationId).toBe(metrics.createCorrelationId("private-turn-id"));
  });

  test("records structural truncation without retaining text", () => {
    const metrics = new BridgeObservability();
    const options = {
      temporaryChat: true,
      defaultTimeZone: "Europe/Stockholm",
      substrate: {
        truncateBeforeSending: true,
        maxSendChars: 400,
        source: "officeweb",
        optionsSets: [],
        allowedMessageTypes: [],
        entityAnnotationTypes: [],
      },
    } as WrapperOptions;
    buildInvocationPayload(
      {
        model: "gpt-5.6",
        promptText: "current user turn",
        additionalContext: [
          { description: "history", text: "x".repeat(1_000) },
        ],
      } as never,
      "conversation",
      "session",
      "request",
      true,
      options,
      metrics,
    );
    const readiness = metrics.readiness();
    const recent = readiness.recentEvents as JsonObject[];
    const event = recent.find((candidate) => candidate.name === "truncation");
    expect(event).toBeDefined();
    expect(JSON.stringify(event)).not.toContain("current user turn");
    expect(JSON.stringify(event)).not.toContain("history");
  });

  test("records bounded store eviction without identifiers", () => {
    const metrics = new BridgeObservability();
    const options = { conversationTtlMinutes: 180 } as WrapperOptions;
    const store = new ConversationStore(options, new DurableStateStore(), metrics);
    for (let index = 0; index <= 1_024; index += 1) {
      store.set("private-key-" + index, "private-conversation-" + index);
    }
    const readiness = metrics.readiness();
    const recent = readiness.recentEvents as JsonObject[];
    const event = recent.find((candidate) => candidate.name === "store_eviction");
    expect(event).toBeDefined();
    expect((event?.fields as JsonObject).reason).toBe("lru");
    expect(JSON.stringify(event)).not.toContain("private-key");
    expect(JSON.stringify(event)).not.toContain("private-conversation");
  });
});
