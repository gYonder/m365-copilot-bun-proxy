import { describe, expect, test } from "bun:test";
import {
  mkdtemp,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
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

  test("retains sanitized protocol sizes without retaining prompt content", () => {
    const metrics = new BridgeObservability();
    const event = metrics.record("retry", {
      reason: "simulated_protocol_correction",
      rejectionReason: "malformed_json",
      requestChars: 83_838,
      assistantTextSize: 13_776,
      prompt: "private prompt",
    });

    expect(event.fields.requestChars).toBe(83_838);
    expect(event.fields.assistantTextSize).toBe(13_776);
    expect(event.fields.prompt).toBe("[redacted]");
    expect(JSON.stringify(event)).not.toContain("private prompt");
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

  test("persists sanitized events to a private rotating JSONL log", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "bridge-events-"));
    const logPath = path.join(directory, "proxy-events.jsonl");
    try {
      const metrics = new BridgeObservability({
        enabled: true,
        logPath,
        maxBytes: 260,
        maxFiles: 2,
      });
      for (let index = 0; index < 6; index += 1) {
        metrics.record("retry", {
          reason: "simulated_protocol_correction",
          attempt: index,
          prompt: `private-${index}`,
        });
      }

      const files = (await readdir(directory)).sort();
      expect(files).toEqual(["proxy-events.jsonl", "proxy-events.jsonl.1"]);
      const persisted = await Promise.all(
        files.map((file) => readFile(path.join(directory, file), "utf8")),
      );
      expect(persisted.join("")).not.toContain("private-");
      expect(persisted.join("")).toContain('"prompt":"[redacted]"');
      expect((await stat(logPath)).mode & 0o777).toBe(0o600);
      expect((metrics.readiness().eventLog as JsonObject).healthy).toBeTrue();

      await writeFile(`${logPath}.2`, "stale\n");
      await writeFile(`${logPath}.3`, "stale\n");
      metrics.record("retry", {
        reason: "oversized",
        detail: "x".repeat(2_000),
      });
      const rotatedFiles = (await readdir(directory)).sort();
      expect(rotatedFiles).toEqual([
        "proxy-events.jsonl",
        "proxy-events.jsonl.1",
      ]);
      for (const file of rotatedFiles) {
        expect((await stat(path.join(directory, file))).size).toBeLessThanOrEqual(
          260,
        );
      }
      const current = await readFile(logPath, "utf8");
      expect(current).toContain('"truncated":true');
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  test("does not create an event log when persistence is disabled", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "bridge-events-"));
    const logPath = path.join(directory, "proxy-events.jsonl");
    try {
      const metrics = new BridgeObservability({
        enabled: false,
        logPath,
        maxBytes: 1_024,
        maxFiles: 2,
      });
      metrics.record("retry", { reason: "disabled" });

      expect(await readdir(directory)).toEqual([]);
      expect((metrics.readiness().eventLog as JsonObject).enabled).toBeFalse();
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});
