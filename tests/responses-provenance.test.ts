import { describe, expect, test } from "bun:test";
import { buildInvocationPayload } from "../src/proxy/clients";
import {
  decodeSubstrateUrlCitations,
  IncrementalPrivateCitationMarkerSanitizer,
  stripPrivateCitationMarkers,
} from "../src/proxy/responses-provenance";
import { deriveRequestProfile } from "../src/proxy/substrate-profiles";
import {
  OpenAiTransformModes,
  ToolChoiceModes,
  TransportNames,
  type JsonObject,
  type ParsedOpenAiRequest,
  type WrapperOptions,
} from "../src/proxy/types";

describe("M365 native search compatibility", () => {
  test("omits Bing for an ordinary coding request", () => {
    const payload = buildInvocationPayload(
      makeRequest(),
      "conversation-1",
      "session-1",
      "request-1",
      true,
      createOptions(),
    );

    expect((payload.arguments[0] as JsonObject).plugins).toEqual([]);
  });

  test("advertises Bing when the prompt asks for current web information", () => {
    const payload = buildInvocationPayload(
      {
        ...makeRequest(),
        hostedWebSearch: true,
        promptText: "What is the current top headline on BBC News?",
      },
      "conversation-1",
      "session-1",
      "request-1",
      true,
      createOptions(),
    );

    expect((payload.arguments[0] as JsonObject).plugins).toEqual([
      { Id: "BingWebSearch", Source: "BuiltIn" },
    ]);
  });

  test("advertises Bing when a request explicitly signals hosted web search", () => {
    const profile = deriveRequestProfile({
      requestJson: {
        model: "m365-copilot",
        input: "inspect",
        tools: [{ type: "web_search" }],
      },
      request: makeRequest(),
      options: createOptions(),
      transport: TransportNames.Substrate,
    });
    expect(profile.hostedWebSearch).toBeTrue();

    const payload = buildInvocationPayload(
      { ...makeRequest(), hostedWebSearch: profile.hostedWebSearch },
      "conversation-1",
      "session-1",
      "request-1",
      true,
      createOptions(),
    );

    expect((payload.arguments[0] as JsonObject).plugins).toEqual([
      { Id: "BingWebSearch", Source: "BuiltIn" },
    ]);
  });

  test("omits Bing when an initial local action is required", () => {
    const request = makeRequest();
    const payload = buildInvocationPayload(
      {
        ...request,
        promptText: "Search the local repository and run its tests.",
        tooling: { ...request.tooling, requiredByLocalAction: true },
      },
      "conversation-1",
      "session-1",
      "request-1",
      true,
      createOptions(),
    );

    expect((payload.arguments[0] as JsonObject).plugins).toEqual([]);
  });

  test("omits Bing when web search is turned off in config", () => {
    const options = createOptions();
    const payload = buildInvocationPayload(
      makeRequest(),
      "conversation-1",
      "session-1",
      "request-1",
      true,
      { ...options, substrate: { ...options.substrate, webSearch: "off" } },
    );

    expect((payload.arguments[0] as JsonObject).plugins).toEqual([]);
  });
});

describe("private citation marker handling", () => {
  test("removes an embedded marker and tidies the sentence", () => {
    expect(
      stripPrivateCitationMarkers(
        "The answer citeturn1search1turn1search3 is ready.",
      ),
    ).toBe("The answer is ready.");
  });

  test("preserves real markdown links", () => {
    const text = "Read [the source](https://example.com/source).";
    expect(stripPrivateCitationMarkers(text)).toBe(text);
  });

  test("returns marker-free text byte-identically", () => {
    const text = "A  line with\textra spacing.";
    expect(stripPrivateCitationMarkers(text)).toBe(text);
  });

  test("removes a marker split across chunk boundaries", () => {
    const sanitizer = new IncrementalPrivateCitationMarkerSanitizer();
    const output = [
      sanitizer.push("See "),
      sanitizer.push("citeturn1sea"),
      sanitizer.push("rch1 more"),
      sanitizer.finish(),
    ].join("");

    expect(output).toBe("See more");
  });

  test("matches whole-string sanitization across streamed chunks", () => {
    const inputs = [
      "Markerless text stays unchanged.",
      "See citeturn1search1 more.",
      "Before citeturn1search1 after.",
      "A longer markerless response with ordinary spacing.",
    ];

    for (const input of inputs) {
      const sanitizer = new IncrementalPrivateCitationMarkerSanitizer();
      const streamed = [
        sanitizer.push(input.slice(0, 5)),
        sanitizer.push(input.slice(5, 13)),
        sanitizer.push(input.slice(13)),
        sanitizer.finish(),
      ].join("");
      expect(streamed).toBe(stripPrivateCitationMarkers(input));
    }
  });

  test("releases text when the bounded holdback is exceeded", () => {
    const sanitizer = new IncrementalPrivateCitationMarkerSanitizer();
    const first = sanitizer.push(`${"a".repeat(200)}c`);
    const rest = sanitizer.finish();

    expect(first.length).toBeGreaterThan(100);
    expect(first + rest).toBe(`${"a".repeat(200)}c`);
  });

  test("holds a long partial citation marker instead of leaking it", () => {
    const sanitizer = new IncrementalPrivateCitationMarkerSanitizer();
    const marker = `cite${"turn1search1".repeat(10)}`;
    const first = sanitizer.push(`Before ${marker.slice(0, 80)}`);
    const second = sanitizer.push(marker.slice(80));

    expect(first).toBe("Before");
    expect(second).toBe("");
    expect(sanitizer.finish()).toBe("");
  });
});

describe("defensive substrate citation decoding", () => {
  test("emits nothing when source data is absent or malformed", () => {
    expect(
      decodeSubstrateUrlCitations(null, "The answer."),
    ).toEqual([]);
    expect(
      decodeSubstrateUrlCitations(
        {
          frames: [
            {
              messageType: "SemanticSerp",
              references: [{ referenceId: "bad", url: "not-a-url" }],
            },
          ],
        },
        "The answer.",
      ),
    ).toEqual([]);
  });

  test("emits one deduplicated citation for a well-formed source", () => {
    const source = {
      referenceId: "ref-1",
      url: "https://example.com/source",
      title: "Example source",
    };
    const citations = decodeSubstrateUrlCitations(
      {
        streamType: "signalr",
        frames: [
          {
            type: 1,
            arguments: [
              {
                messages: [
                  { messageType: "SemanticSerp", references: [source] },
                  {
                    messageType: "ReferencesListComplete",
                    references: [source],
                  },
                ],
              },
            ],
          },
        ],
      },
      "Read the Example source for details.",
    );

    expect(citations).toEqual([
      {
        type: "url_citation",
        start_index: 9,
        end_index: 23,
        url: source.url,
        title: source.title,
      },
    ]);
  });
});

function makeRequest(): ParsedOpenAiRequest {
  return {
    model: "m365-copilot",
    stream: false,
    transformMode: OpenAiTransformModes.Mapped,
    promptText: "inspect",
    userKey: null,
    locationHint: {},
    contextualResources: null,
    additionalContext: [],
    tooling: {
      tools: [],
      toolChoiceMode: ToolChoiceModes.Auto,
      toolChoiceFunctionName: null,
      parallelToolCalls: false,
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
      isEdu: false,
      agent: "web",
      variants: null,
      clientPlatform: "mcmcopilot-web",
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
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
    logStdout: false,
    confabRetries: 1,
    msalAuth: true,
    imageGeneration: {
      enabled: false,
      maxPromptChars: 4_000,
      maxImages: 1,
      maxArtifactBytes: 20_000_000,
      timeoutMs: 120_000,
      concurrencyLimit: 1,
      allowedMimeTypes: ["image/png"],
    },
  };
}
