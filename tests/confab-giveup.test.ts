import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import { looksLikeConfabGiveUp } from "../src/proxy/server";
import { createProxyApp } from "../src/proxy/server";
import { ResponseStore } from "../src/proxy/response-store";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import {
  LogLevels,
  OpenAiTransformModes,
  PlaywrightBrowsers,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";
import { BridgeObservability } from "../src/proxy/observability";

describe("looksLikeConfabGiveUp", () => {
  test("detects the canonical give-up phrasings", () => {
    const samples = [
      "It seems the file-editing tools are not available right now. Please restart the task.",
      "The file editing tools are currently not available.",
      "I don't have access to the file-editing tools needed for this.",
      "I'm unable to use the code tools, please try to restart the task.",
    ];
    for (const sample of samples) {
      expect(looksLikeConfabGiveUp(sample)).toBeTrue();
    }
  });

  test("does not flag a normal successful answer", () => {
    const samples = [
      "",
      "   ",
      "Here is the refactored function you asked for.",
      "I edited the file and updated the tests accordingly.",
      "The available tools worked and the task is complete.",
    ];
    for (const sample of samples) {
      expect(looksLikeConfabGiveUp(sample)).toBeFalse();
    }
  });
});

describe("confab recovery terminal outcomes", () => {
  test("failed re-ask becomes ambiguous_completion", async () => {
    const calls: string[] = [];
    const observability = new BridgeObservability();
    const app = createConfabApp(
      [
        buildSuccessResult(
          "The file-editing tools are not available. Please restart the task.",
        ),
        {
          isSuccess: false,
          statusCode: 502,
          responseJson: null,
          rawBody: JSON.stringify({
            message: "upstream failed",
            code: "transport_failed",
          }),
          assistantText: null,
          conversationId: null,
          errorCode: "transport_failed",
        },
      ],
      calls,
      observability,
    );

    const response = await postChat(app);
    const body = (await response.json()) as JsonObject;

    expect(response.status).toBe(502);
    expect((body.error as JsonObject).code).toBe("ambiguous_completion");
    expect(calls).toHaveLength(2);
    expect(calls[0]).not.toBe(calls[1]);
    const ready = await app.fetch(new Request("http://localhost/readyz"));
    expect(JSON.stringify(await ready.json())).toContain("confab_give_up");
  });

  test("successful re-ask returns the real answer", async () => {
    const calls: string[] = [];
    const app = createConfabApp(
      [
        buildSuccessResult(
          "I can't use the code tools, please try to restart the task.",
        ),
        buildSuccessResult("The real answer from the successful re-ask."),
      ],
      calls,
    );

    const response = await postChat(app);
    const body = (await response.json()) as JsonObject;
    const choices = body.choices as JsonObject[];
    const message = choices[0]?.message as JsonObject;

    expect(response.status).toBe(200);
    expect(message.content).toBe("The real answer from the successful re-ask.");
    expect(calls).toHaveLength(2);
    expect(calls[0]).not.toBe(calls[1]);
  });
});

function createConfabApp(
  results: ChatResult[],
  calls: string[],
  observability?: BridgeObservability,
) {
  const options = createOptions();
  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "graph-unused",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("Graph chat is not used in confab tests.");
    },
    chatOverStream: async (): Promise<Response> => {
      throw new Error("Graph streaming is not used in confab tests.");
    },
  } as unknown as CopilotGraphClient;
  let conversationSequence = 0;
  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: `substrate-${++conversationSequence}`,
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
    ): Promise<ChatResult> => {
      calls.push(conversationId);
      const result = results.shift();
      if (!result) {
        throw new Error("Unexpected extra confab test call.");
      }
      return result;
    },
    chatStream: async (): Promise<ChatResult> => {
      throw new Error("Substrate streaming is not used in confab tests.");
    },
  } as unknown as CopilotSubstrateClient;
  const debugLogger = {
    logIncomingRequest: async () => {},
    logOutgoingResponse: async () => {},
    logUpstreamRequest: async () => {},
    logUpstreamResponse: async () => {},
    logSubstrateFrame: async () => {},
  } as unknown as DebugMarkdownLogger;
  const tokenProvider = {
    resolveAuthorizationHeader: async () => "******",
  } as unknown as ProxyTokenProvider;

  return createProxyApp({
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore: new ConversationStore(options),
    responseStore: new ResponseStore(options),
    tokenProvider,
    observability,
  });
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Mapped,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: PlaywrightBrowsers.Edge,
    transport: TransportNames.Substrate,
    graphBaseUrl: "https://graph.microsoft.com",
    createConversationPath: "/beta/copilot/conversations",
    chatPathTemplate: "/beta/copilot/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate:
      "/beta/copilot/conversations/{conversationId}/chatOverStream",
    substrate: {
      hubPath: "wss://substrate.office.com/m365Copilot/Chathub",
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
    logStdout: false,
    confabRetries: 1,
    msalAuth: false,
    imageGeneration: {
      enabled: false,
      maxPromptChars: 1_000,
      maxImages: 1,
      maxArtifactBytes: 1_000,
      timeoutMs: 1_000,
      concurrencyLimit: 1,
      allowedMimeTypes: ["image/png"],
    },
  };
}

function buildSuccessResult(assistantText: string): ChatResult {
  return {
    isSuccess: true,
    statusCode: 200,
    responseJson: null,
    rawBody: "",
    assistantText,
    conversationId: "upstream-conversation",
  };
}

async function postChat(
  app: ReturnType<typeof createProxyApp>,
): Promise<Response> {
  return app.fetch(
    new Request("http://localhost/v1/chat/completions", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        model: "gpt-5.6-sol",
        stream: false,
        messages: [{ role: "user", content: "Use the tools and answer." }],
      }),
    }),
  );
}
