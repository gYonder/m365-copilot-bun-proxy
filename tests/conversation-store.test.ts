import { describe, expect, test } from "bun:test";
import { ConversationStore } from "../src/proxy/conversation-store";
import type { WrapperOptions } from "../src/proxy/types";

describe("ConversationStore", () => {
  test("evicts the least-recently-used conversation after the cap", () => {
    const store = new ConversationStore(createOptions());
    for (let index = 0; index < 1_024; index += 1) {
      store.set(`key-${index}`, `conversation-${index}`);
    }
    expect(store.tryGet("key-0")).toBe("conversation-0");

    store.set("key-new", "conversation-new");

    expect(store.tryGet("key-1")).toBeNull();
    expect(store.tryGet("key-0")).toBe("conversation-0");
    expect(store.tryGet("key-new")).toBe("conversation-new");
  });
});

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
