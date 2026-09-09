import { describe, expect, test } from "bun:test";
import { DurableStateStore } from "../src/proxy/durable-state";
import { ResponseStore } from "../src/proxy/response-store";
import { computeResponsesReplayIdentityKey } from "../src/proxy/server";
import type {
  JsonObject,
  ResponsesProtocolIdentity,
  WrapperOptions,
} from "../src/proxy/types";

describe("ResponseStore replay identity", () => {
  test("binds each request hash to its exact completed response", () => {
    const store = new ResponseStore(createOptions());
    const first = response("resp-first", "first");
    const second = response("resp-second", "second");

    store.rememberCompletedRequest("hash-first", "conversation-1", first);
    store.rememberCompletedRequest("hash-second", "conversation-1", second);

    expect(store.tryGetRequestReplay("hash-first")?.response).toEqual(first);
    expect(store.tryGetRequestReplay("hash-second")?.response).toEqual(second);
  });

  test("does not claim a request before a completed response is stored", () => {
    const store = new ResponseStore(createOptions());
    expect(store.tryGetRequestReplay("in-flight-hash")).toBeNull();
  });

  test("shares one in-flight response promise by protocol identity", async () => {
    const store = new ResponseStore(createOptions());
    const responsePromise = Promise.resolve(new Response("ok"));

    expect(
      store.registerInFlightResponse("protocol:turn-1", responsePromise),
    ).toBe(responsePromise);
    expect(store.tryGetInFlightResponse("protocol:turn-1")).toBe(
      responsePromise,
    );
    expect(
      store.registerInFlightResponse(
        "protocol:turn-1",
        Promise.resolve(new Response("duplicate")),
      ),
    ).toBe(responsePromise);

    await responsePromise;
    await Promise.resolve();
    expect(store.tryGetInFlightResponse("protocol:turn-1")).toBeNull();
  });

  test("keeps identical prompts in distinct protocol turns independent", () => {
    const first = computeResponsesReplayIdentityKey(
      identity({ turnId: "turn-1" }),
      { input: "same prompt" },
      "graph",
      "same-request-hash",
    );
    const second = computeResponsesReplayIdentityKey(
      identity({ turnId: "turn-2" }),
      { input: "same prompt" },
      "graph",
      "same-request-hash",
    );

    expect(first).not.toBe(second);
    expect(first.startsWith("protocol:")).toBeTrue();
    expect(second.startsWith("protocol:")).toBeTrue();
  });

  test("canonicalizes multiple tool results regardless of input ordering", () => {
    const protocolIdentity = identity({
      previousResponseId: "resp_previous",
      callIds: ["call_a", "call_b"],
    });
    const first = computeResponsesReplayIdentityKey(
      protocolIdentity,
      {
        input: [
          { type: "function_call_output", call_id: "call_a", output: "one" },
          { type: "custom_tool_call_output", call_id: "call_b", output: "two" },
        ],
      },
      "graph",
      "unused",
    );
    const second = computeResponsesReplayIdentityKey(
      protocolIdentity,
      {
        input: [
          { type: "custom_tool_call_output", call_id: "call_b", output: "two" },
          { type: "function_call_output", call_id: "call_a", output: "one" },
        ],
      },
      "graph",
      "unused",
    );

    expect(first).toBe(second);
  });

  test("uses the bounded legacy request hash when protocol identity is absent", () => {
    expect(
      computeResponsesReplayIdentityKey(
        identity(),
        { input: "legacy prompt" },
        "graph",
        "legacy-request-hash",
      ),
    ).toBe("legacy:legacy-request-hash");
  });

  test("evicts the oldest stored responses after the cap", () => {
    const store = new ResponseStore(createOptions());
    for (let index = 0; index < 1_025; index += 1) {
      const body = response(`resp-${index}`, String(index));
      store.set(`resp-${index}`, body, `conversation-${index}`);
    }

    expect(store.tryGet("resp-0")).toBeNull();
    expect(store.tryGet("resp-1024")?.id).toBe("resp-1024");
  });

  test("stores context usage with its response chain identity", () => {
    const store = new ResponseStore(createOptions());
    store.set("resp-context", response("resp-context", "ok"), "conversation-1", null, 42_000, "window-1");
    expect(store.tryGetContextUsage("resp-context")).toEqual({
      inputTokens: 42_000,
      windowId: "window-1",
    });
  });

  test("keeps transient response links available in memory only", () => {
    const durable = new DurableStateStore();
    const store = new ResponseStore(createOptions(), durable);
    const body = failedResponse("temporary failure");

    store.setTransient("resp-transient", body, "conversation-1");

    expect(store.tryGet("resp-transient")).toEqual(body);
    expect(store.tryGetConversationLink("resp-transient")).toBe(
      "conversation-1",
    );
    expect(durable.state.responses["resp-transient"]).toBeUndefined();
  });

  test("allows one fresh protocol retry before replaying the failure", () => {
    const store = new ResponseStore(createOptions());
    const firstFailure = failedResponse("first failure");
    const secondFailure = failedResponse("second failure");

    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "primary",
    });

    store.rememberRetryableProtocolFailure(
      "protocol:turn-1",
      "conversation-1",
      firstFailure,
    );
    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "regenerate",
      conversationId: "conversation-1",
      response: firstFailure,
    });
    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "replay",
      conversationId: "conversation-1",
      response: firstFailure,
    });

    store.rememberExhaustedProtocolFailure(
      "protocol:turn-1",
      "conversation-2",
      secondFailure,
    );
    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "replay",
      conversationId: "conversation-2",
      response: secondFailure,
    });
  });

  test("clears protocol failure state after a successful completion", () => {
    const store = new ResponseStore(createOptions());
    store.rememberRetryableProtocolFailure(
      "protocol:turn-1",
      null,
      failedResponse("failure"),
    );
    expect(store.claimProtocolFailureCycle("protocol:turn-1").kind).toBe(
      "regenerate",
    );

    store.clearProtocolFailureCycle("protocol:turn-1");

    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "primary",
    });
  });

  test("keeps protocol failure cycles in memory only", () => {
    const firstStore = new ResponseStore(createOptions());
    firstStore.rememberRetryableProtocolFailure(
      "protocol:turn-1",
      null,
      failedResponse("failure"),
    );

    const restartedStore = new ResponseStore(createOptions());

    expect(restartedStore.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "primary",
    });
  });

  test("expires protocol failure cycles without refreshing replay reads", () => {
    let now = 1_000;
    const store = new ResponseStore(createOptions(), undefined, null, () => now);
    store.rememberRetryableProtocolFailure(
      "protocol:turn-1",
      null,
      failedResponse("failure"),
    );
    expect(store.claimProtocolFailureCycle("protocol:turn-1").kind).toBe(
      "regenerate",
    );

    now += 59_000;
    expect(store.claimProtocolFailureCycle("protocol:turn-1").kind).toBe(
      "replay",
    );

    now += 1_001;
    expect(store.claimProtocolFailureCycle("protocol:turn-1")).toEqual({
      kind: "primary",
    });
  });

  test("resolves call IDs only within the request-correlated task scope", () => {
    const store = new ResponseStore(createOptions());
    const ledger1 = store.getOrCreateToolLedger("task-1");
    ledger1.issueCalls({
      taskId: "task-1",
      responseId: "resp-1",
      requestProfileKey: "prof-1",
      calls: [
        { call_id: "call-1", name: "test", type: "function", arguments: {} },
        { call_id: "call-2", name: "test", type: "function", arguments: {} },
      ],
      round: 1,
    });
    store.saveToolLedger("task-1", ledger1);

    const ledger2 = store.getOrCreateToolLedger("task-2");
    ledger2.issueCalls({
      taskId: "task-2",
      responseId: "resp-2",
      requestProfileKey: "prof-1",
      calls: [
        { call_id: "call-3", name: "test", type: "function", arguments: {} },
      ],
      round: 1,
    });
    store.saveToolLedger("task-2", ledger2);

    expect(
      store.resolveToolLedgerTaskScope(["call-1"], "resp-1", "fallback"),
    ).toEqual({
      kind: "resolved",
      taskId: "task-1",
    });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-1", "call-2"],
        "resp-1",
        "fallback",
      ),
    ).toEqual({
      kind: "resolved",
      taskId: "task-1",
    });
    expect(
      store.resolveToolLedgerTaskScope(["call-3"], "resp-2", "fallback"),
    ).toEqual({
      kind: "resolved",
      taskId: "task-2",
    });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-unknown"],
        "resp-1",
        "fallback",
      ),
    ).toEqual({
      kind: "conflict",
    });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-1", "call-unknown"],
        "resp-1",
        "fallback",
      ),
    ).toEqual({
      kind: "conflict",
    });
    expect(
      store.resolveToolLedgerTaskScope(["call-1", "call-3"], "resp-1", "fallback"),
    ).toEqual({
      kind: "conflict",
    });
    expect(
      store.resolveToolLedgerTaskScope(["call-3"], null, "task-2"),
    ).toEqual({
      kind: "resolved",
      taskId: "task-2",
    });
    expect(
      store.resolveToolLedgerTaskScope(["call-unknown"], null, "unknown-task"),
    ).toEqual({ kind: "none" });
    expect(store.resolveToolLedgerTaskScope([], null, "task-1")).toEqual({
      kind: "none",
    });
  });

  test("does not correlate reused call IDs across independent responses", () => {
    const store = new ResponseStore(createOptions());
    for (const [taskId, responseId] of [
      ["task-1", "resp-1"],
      ["task-2", "resp-2"],
    ] as const) {
      const ledger = store.getOrCreateToolLedger(taskId);
      ledger.issueCalls({
        taskId,
        responseId,
        requestProfileKey: "prof-1",
        calls: [
          { call_id: "call-shared", name: "test", type: "function", arguments: {} },
        ],
        round: 1,
      });
      store.saveToolLedger(taskId, ledger);
    }

    expect(
      store.resolveToolLedgerTaskScope(
        ["call-shared"],
        "resp-1",
        "unrelated-task",
      ),
    ).toEqual({
      kind: "resolved",
      taskId: "task-1",
    });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-shared"],
        "resp-unknown",
        "unrelated-task",
      ),
    ).toEqual({ kind: "none" });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-shared"],
        "resp-1",
        "task-2",
      ),
    ).toEqual({
      kind: "resolved",
      taskId: "task-1",
    });
    expect(
      store.resolveToolLedgerTaskScope(
        ["call-3"],
        "resp-1",
        "task-2",
      ),
    ).toEqual({ kind: "conflict" });
  });
});

describe("ResponseStore task deadlines", () => {
  test("replaces an expired deadline during its retention grace period", async () => {
    const store = new ResponseStore(createOptions());
    const expired = store.getOrCreateTaskDeadline(
      "task-1",
      () => Date.now() + 5,
    );
    await Bun.sleep(10);

    const refreshed = store.getOrCreateTaskDeadline(
      "task-1",
      () => Date.now() + 1_000,
    );

    expect(refreshed).toBeGreaterThan(expired);
    expect(refreshed).toBeGreaterThan(Date.now());
  });
});

function response(id: string, text: string): JsonObject {
  return {
    id,
    object: "response",
    status: "completed",
    output: [
      {
        type: "message",
        role: "assistant",
        content: [{ type: "output_text", text }],
      },
    ],
    output_text: text,
  };
}

function failedResponse(message: string): JsonObject {
  return {
    id: "resp-failed",
    object: "response",
    status: "failed",
    error: {
      code: "provider_drift",
      message,
    },
  };
}

function identity(
  overrides: Partial<ResponsesProtocolIdentity> = {},
): ResponsesProtocolIdentity {
  return {
    threadId: null,
    sessionId: null,
    turnId: null,
    conversationId: null,
    previousResponseId: null,
    callIds: [],
    ...overrides,
  };
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: "error",
    openAiTransformMode: "simulated",
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: "substrate",
    graphBaseUrl: "https://graph.microsoft.com",
    createConversationPath: "/conversations",
    chatPathTemplate: "/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate: "/conversations/{conversationId}/stream",
    substrate: {
      hubPath: "wss://example.invalid",
      source: "officeweb",
      quoteSourceInQuery: true,
      scenario: "OfficeWebIncludedCopilot",
      origin: "https://m365.cloud.microsoft",
      product: "Office",
      agentHost: "Bizchat.FullScreen",
      licenseType: "Starter",
      agent: "web",
      variants: null,
      clientPlatform: "web",
      productThreadType: "Office",
      invocationTimeoutSeconds: 120,
      taskTimeoutSeconds: 900,
      keepAliveSeconds: 15,
      optionsSets: [],
      allowedMessageTypes: [],
      invocationTarget: "chat",
      invocationType: 4,
      locale: "en-US",
      experienceType: "Default",
      entityAnnotationTypes: [],
      earlyCompleteOnSimulatedPayload: false,
    },
    defaultModel: "gpt-5.6-sol",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}
