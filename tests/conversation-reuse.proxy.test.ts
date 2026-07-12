import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import { ResponseStore } from "../src/proxy/response-store";
import { createProxyApp } from "../src/proxy/server";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";

describe("conversation reuse", () => {
  test("reuses conversation id across follow-up calls when user key is stable", async () => {
    const options = createOptions();
    const conversationStore = new ConversationStore(options);
    const responseStore = new ResponseStore(options);
    let createConversationCount = 0;

    const graphClient = {
      createConversation: async (): Promise<CreateConversationResult> => ({
        isSuccess: true,
        statusCode: 200,
        conversationId: `conv-${++createConversationCount}`,
        rawBody: "{}",
      }),
      chat: async (
        _authorizationHeader: string,
        conversationId: string,
        _payload: JsonObject,
      ): Promise<ChatResult> => ({
        isSuccess: true,
        statusCode: 200,
        responseJson: {
          id: conversationId,
          messages: [{ text: "user" }, { text: "assistant" }],
        },
        rawBody: "{}",
        assistantText: "assistant",
        conversationId,
      }),
      chatOverStream: async (): Promise<Response> => {
        throw new Error("not used");
      },
    } as unknown as CopilotGraphClient;

    const substrateClient = {
      createConversation: (): CreateConversationResult => ({
        isSuccess: true,
        statusCode: 200,
        conversationId: "conv-substrate",
        rawBody: "{}",
      }),
      chat: async (): Promise<ChatResult> => {
        throw new Error("not used");
      },
      chatStream: async (): Promise<ChatResult> => {
        throw new Error("not used");
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
      resolveAuthorizationHeader: async () => "Bearer unit-test-token",
    } as unknown as ProxyTokenProvider;

    const app = createProxyApp({
      options,
      debugLogger,
      graphClient,
      substrateClient,
      conversationStore,
      responseStore,
      tokenProvider,
    });

    const headers = {
      "content-type": "application/json",
      "x-m365-transport": TransportNames.Graph,
    };
    const first = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers,
        body: JSON.stringify({
          model: "m365-copilot",
          user: "roo-thread-1",
          messages: [{ role: "user", content: "hello" }],
        }),
      }),
    );
    const second = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers,
        body: JSON.stringify({
          model: "m365-copilot",
          user: "roo-thread-1",
          messages: [{ role: "user", content: "follow up" }],
        }),
      }),
    );

    expect(first.status).toBe(200);
    expect(second.status).toBe(200);
    expect(first.headers.get("x-m365-conversation-id")).toBe("conv-1");
    expect(second.headers.get("x-m365-conversation-id")).toBe("conv-1");
    expect(first.headers.get("x-m365-conversation-created")).toBe("true");
    expect(second.headers.get("x-m365-conversation-created")).toBeNull();
    expect(createConversationCount).toBe(1);
  });

  test("allows an identical retry after an aborted Responses stream before upstream completion", async () => {
    const options = createOptions();
    const conversationStore = new ConversationStore(options);
    const responseStore = new ResponseStore(options);
    let upstreamCallCount = 0;
    const releaseUpstream: Array<() => void> = [];

    const graphClient = {
      createConversation: async (): Promise<CreateConversationResult> => ({
        isSuccess: true,
        statusCode: 200,
        conversationId: "conv-stream-race",
        rawBody: "{}",
      }),
      chat: async (): Promise<ChatResult> => {
        throw new Error("not used");
      },
      chatOverStream: async (): Promise<Response> => {
        const callNumber = ++upstreamCallCount;
        const stream = new ReadableStream<Uint8Array>({
          start(controller) {
            let closed = false;
            controller.enqueue(
              new TextEncoder().encode(
                `data: ${JSON.stringify({
                  id: "conv-stream-race",
                  messages: [{ text: `reply-${callNumber}` }],
                })}\n\n`,
              ),
            );
            releaseUpstream[callNumber - 1] = () => {
              if (!closed) {
                closed = true;
                controller.close();
              }
            };
          },
        });
        return new Response(stream, {
          status: 200,
          headers: { "content-type": "text/event-stream" },
        });
      },
    } as unknown as CopilotGraphClient;

    const substrateClient = {
      createConversation: (): CreateConversationResult => ({
        isSuccess: true,
        statusCode: 200,
        conversationId: "conv-substrate",
        rawBody: "{}",
      }),
      chat: async (): Promise<ChatResult> => {
        throw new Error("not used");
      },
      chatStream: async (): Promise<ChatResult> => {
        throw new Error("not used");
      },
    } as unknown as CopilotSubstrateClient;

    const debugLogger = {
      logIncomingRequest: async () => {},
      logOutgoingResponse: async () => {},
      logUpstreamRequest: async () => {},
      logUpstreamResponse: async () => {},
      logSubstrateFrame: async () => {},
      logOutgoingStreamBody: async () => {},
    } as unknown as DebugMarkdownLogger;

    const tokenProvider = {
      resolveAuthorizationHeader: async () => "******",
    } as unknown as ProxyTokenProvider;

    const app = createProxyApp({
      options,
      debugLogger,
      graphClient,
      substrateClient,
      conversationStore,
      responseStore,
      tokenProvider,
    });
    const server = Bun.serve({
      port: 0,
      fetch: app.fetch,
    });

    const requestBody = JSON.stringify({
      model: "m365-copilot",
      stream: true,
      input: "hello",
    });
    const headers = {
      "content-type": "application/json",
      "x-m365-transport": TransportNames.Graph,
    };
    const firstAbort = new AbortController();

    try {
      const first = await fetch(`http://localhost:${server.port}/v1/responses`, {
        method: "POST",
        headers,
        body: requestBody,
        signal: firstAbort.signal,
      });
      expect(first.status).toBe(200);
      expect(first.body).not.toBeNull();

      const firstReader = first.body?.getReader();
      expect(firstReader).toBeDefined();
      let firstText = "";
      while (!firstText.includes('"delta":"reply-1"')) {
        const next = await firstReader?.read();
        if (!next || next.done) {
          break;
        }
        firstText += new TextDecoder().decode(next.value);
      }
      expect(firstText).toContain('"delta":"reply-1"');
      firstAbort.abort();
      await firstReader?.cancel();

      const retryPromise = fetch(`http://localhost:${server.port}/v1/responses`, {
        method: "POST",
        headers,
        body: requestBody,
      });
      while (upstreamCallCount < 2) {
        await Bun.sleep(1);
      }
      releaseUpstream[1]?.();

      const retry = await retryPromise;
      expect(retry.status).toBe(200);
      expect(retry.status).not.toBe(409);
      await retry.text();
      expect(upstreamCallCount).toBe(2);
    } finally {
      releaseUpstream[0]?.();
      releaseUpstream[1]?.();
      server.stop(true);
    }
  });
});

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Mapped,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Graph,
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
      earlyCompleteOnSimulatedPayload: false,
      entityAnnotationTypes: [],
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}
