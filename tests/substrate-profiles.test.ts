import { describe, expect, test } from "bun:test";
import { deriveRequestProfile, isSupportedTaskScope } from "../src/proxy/substrate-profiles";
import {
  OpenAiTransformModes,
  TransportNames,
  type OpenAiToolDefinition,
  type ParsedOpenAiRequest,
  type WrapperOptions,
} from "../src/proxy/types";

describe("substrate request profiles", () => {
  test("canonicalizes semantically identical structured requests", () => {
    const options = createOptions();
    const request = createParsedRequest();
    const first = deriveRequestProfile({
      requestJson: {
        model: "m365-copilot",
        tools: [
          {
            type: "function",
            name: "read_file",
            description: "Read a file.",
            parameters: {
              type: "object",
              properties: { path: { type: "string" } },
            },
          },
        ],
        input: [{ type: "message", role: "user", content: "inspect" }],
      },
      request,
      options,
      transport: TransportNames.Substrate,
    });
    const second = deriveRequestProfile({
      requestJson: {
        input: [{ content: "inspect", role: "user", type: "message" }],
        tools: [
          {
            parameters: {
              properties: { path: { type: "string" } },
              type: "object",
            },
            description: "Read a file.",
            name: "read_file",
            type: "function",
          },
        ],
        model: "m365-copilot",
      },
      request,
      options,
      transport: TransportNames.Substrate,
    });

    expect(second.compatibilityKey).toBe(first.compatibilityKey);
    expect(second.profileId).toBe(first.profileId);
  });

  test("changes the compatibility key for capability differences", () => {
    const options = createOptions();
    const request = createParsedRequest();
    const coding = deriveRequestProfile({
      requestJson: { model: "m365-copilot", input: "inspect" },
      request,
      options,
      transport: TransportNames.Substrate,
    });
    const toolEnabled = deriveRequestProfile({
      requestJson: {
        model: "m365-copilot",
        input: "inspect",
        tools: [
          {
            type: "function",
            name: "read_file",
            parameters: { type: "object" },
          },
        ],
      },
      request: {
        ...request,
        tooling: {
          ...request.tooling,
          tools: [
            {
              name: "read_file",
              type: "function",
              description: null,
              parameters: { type: "object" },
              format: null,
            },
          ],
        },
      },
      options,
      transport: TransportNames.Substrate,
    });
    const searchEnabled = deriveRequestProfile({
      requestJson: {
        model: "m365-copilot",
        input: "inspect",
        tools: [{ type: "web_search_preview" }],
      },
      request,
      options,
      transport: TransportNames.Substrate,
    });
    const imageEnabled = deriveRequestProfile({
      requestJson: {
        model: "m365-copilot",
        input: [
          {
            type: "message",
            role: "user",
            content: [{ type: "input_image", image_url: "https://example.invalid/a" }],
          },
        ],
      },
      request,
      options,
      transport: TransportNames.Substrate,
    });

    expect(toolEnabled.compatibilityKey).not.toBe(coding.compatibilityKey);
    expect(searchEnabled.compatibilityKey).not.toBe(coding.compatibilityKey);
    expect(imageEnabled.compatibilityKey).not.toBe(coding.compatibilityKey);
    expect(searchEnabled.hostedWebSearch).toBeTrue();
    expect(imageEnabled.imageInput).toBeTrue();
  });

  test("does not reject a tool schema beyond canonical JSON bounds", () => {
    const request = createParsedRequest();
    const tool: OpenAiToolDefinition = {
      name: "large_schema",
      type: "function",
      description: "x".repeat(70_000),
      parameters: { type: "object" },
      format: null,
    };

    expect(() =>
      deriveRequestProfile({
        requestJson: { model: "m365-copilot", input: "inspect" },
        request: {
          ...request,
          tooling: { ...request.tooling, tools: [tool] },
        },
        options: createOptions(),
        transport: TransportNames.Substrate,
      }),
    ).not.toThrow();
  });
});

function createParsedRequest(): ParsedOpenAiRequest {
  return {
    model: "m365-copilot",
    stream: false,
    transformMode: OpenAiTransformModes.Mapped,
    promptText: "inspect",
    userKey: "profile-test",
    locationHint: {},
    contextualResources: null,
    additionalContext: [],
    tooling: {
      tools: [],
      toolChoiceMode: "none",
      toolChoiceFunctionName: null,
      parallelToolCalls: true,
    },
    responseFormat: null,
    reasoningEffort: null,
    temperature: null,
  };
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: "error",
    openAiTransformMode: OpenAiTransformModes.Mapped,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Substrate,
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
      optionsSets: ["search", "code"],
      allowedMessageTypes: ["Chat"],
      invocationTarget: "chat",
      invocationType: 4,
      locale: "en-US",
      experienceType: "Default",
      entityAnnotationTypes: [],
      earlyCompleteOnSimulatedPayload: false,
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}

describe("task scope aliases", () => {
  test("accepts the canonical scope and its natural short spelling", () => {
    expect(isSupportedTaskScope("coding_project")).toBe(true);
    expect(isSupportedTaskScope("coding")).toBe(true);
    expect(isSupportedTaskScope("Coding")).toBe(true);
    expect(isSupportedTaskScope("")).toBe(true);
  });

  test("still rejects scopes this coding-only provider cannot serve", () => {
    expect(isSupportedTaskScope("chat")).toBe(false);
    expect(isSupportedTaskScope("research")).toBe(false);
  });
});
