import { describe, expect, test } from "bun:test";
import { ResponseStore } from "../src/proxy/response-store";
import type { JsonObject, WrapperOptions } from "../src/proxy/types";

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
