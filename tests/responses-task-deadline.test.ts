import { describe, expect, test } from "bun:test";
import { ResponseStore } from "../src/proxy/response-store";
import {
  computeResponsesTaskKey,
  resolveResponsesTaskDeadlineMs,
} from "../src/proxy/server";
import {
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";

describe("Responses task deadlines", () => {
  test("inherits the original deadline when Codex appends tool activity", () => {
    const options = createOptions();
    const store = new ResponseStore(options);
    const initial = request([]);
    const inheritedDeadline = resolveResponsesTaskDeadlineMs(
      initial,
      store,
      options,
    );

    const deadline = resolveResponsesTaskDeadlineMs(
      request([
        {
          type: "function_call",
          call_id: "call-1",
          name: "exec",
          arguments: "{}",
        },
        {
          type: "function_call_output",
          call_id: "call-1",
          output: "done",
        },
      ]),
      store,
      options,
    );

    expect(deadline).toBe(inheritedDeadline);
  });

  test("starts a fresh budget for a new user turn", () => {
    const options = createOptions();
    const store = new ResponseStore(options);
    const initial = request([]);
    resolveResponsesTaskDeadlineMs(initial, store, options);

    const before = Date.now();
    const deadline = resolveResponsesTaskDeadlineMs(
      request([
        {
          role: "user",
          type: "message",
          content: [{ type: "input_text", text: "next" }],
        },
      ]),
      store,
      options,
    );

    expect(deadline).toBeGreaterThanOrEqual(before + 149_000);
    expect(computeResponsesTaskKey(initial)).not.toBe(
      computeResponsesTaskKey(
        request([
          {
            role: "user",
            type: "message",
            content: [{ type: "input_text", text: "next" }],
          },
        ]),
      ),
    );
  });
});

function request(appendedItems: JsonObject[]): JsonObject {
  return {
    prompt_cache_key: "task-key-1",
    input: [
      {
        role: "user",
        type: "message",
        content: [{ type: "input_text", text: "do work" }],
      },
      ...appendedItems,
    ],
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
      taskTimeoutSeconds: 150,
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
