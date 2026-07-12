import { describe, expect, test } from "bun:test";
import {
  CopilotSubstrateClient,
  type SubstrateReceiverFactory,
  type SubstrateWebSocketConnector,
} from "../src/proxy/clients";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  LogLevels,
  OpenAiTransformModes,
  ToolChoiceModes,
  TransportNames,
  type ParsedOpenAiRequest,
  type WrapperOptions,
} from "../src/proxy/types";

type FakeWebSocket = Parameters<SubstrateReceiverFactory>[0];

const stubLogger = {
  logSubstrateFrame: async () => {},
  logIncomingRequest: async () => {},
  logOutgoingResponse: async () => {},
  logUpstreamRequest: async () => {},
  logUpstreamResponse: async () => {},
} as unknown as DebugMarkdownLogger;

function makeJwtAuthHeader(): string {
  const header = Buffer.from(JSON.stringify({ alg: "none" })).toString(
    "base64url",
  );
  const payload = Buffer.from(
    JSON.stringify({
      oid: "00000000-0000-0000-0000-000000000001",
      tid: "00000000-0000-0000-0000-000000000002",
    }),
  ).toString("base64url");
  return `Bearer ${header}.${payload}.signature`;
}

function makeRequest(): ParsedOpenAiRequest {
  return {
    model: "gpt-5.6-sol",
    stream: false,
    transformMode: OpenAiTransformModes.Simulated,
    promptText: "hello",
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
    reasoningEffort: "high",
    temperature: null,
  };
}

/**
 * Builds a fake Substrate transport whose receiver hands back a valid
 * handshake, then blocks on the first in-turn read. `onBlockingRead` fires at
 * that point so the test can simulate a client disconnect mid-turn. Closing the
 * socket makes subsequent reads resolve to null (mirroring the real receiver's
 * close→flush(null) behaviour), which is how the turn loop unwinds.
 */
function makeFakeTransport(onBlockingRead: () => void): {
  connect: SubstrateWebSocketConnector;
  createReceiver: SubstrateReceiverFactory;
  closeCalls: Array<{ code?: number; reason?: string }>;
} {
  const closeCalls: Array<{ code?: number; reason?: string }> = [];
  let closed = false;
  const ws = {
    readyState: 1, // WebSocket.OPEN
    send: (_payload: string) => {},
    close: (code?: number, reason?: string) => {
      if (!closed) {
        closed = true;
        (ws as { readyState: number }).readyState = 3; // WebSocket.CLOSED
        closeCalls.push({ code, reason });
      }
    },
  };

  let nextCall = 0;
  const receiver = {
    next: async (_timeoutMs: number): Promise<string | null> => {
      nextCall += 1;
      if (closed) {
        return null;
      }
      if (nextCall === 1) {
        return "{}"; // handshake frame, no error
      }
      if (nextCall === 2) {
        // The turn loop is now awaiting upstream output. Simulate the client
        // going away; the abort handler must close the socket.
        onBlockingRead();
      }
      // Once the abort closed the socket, the read completes with null and the
      // loop terminates. Otherwise it would block forever (the real "in-flight"
      // state) — which is exactly what cancellation must interrupt.
      return closed ? null : await new Promise<string | null>(() => {});
    },
    dispose: () => {},
  };

  const connect: SubstrateWebSocketConnector = async () =>
    ws as unknown as FakeWebSocket;
  const createReceiver: SubstrateReceiverFactory = () => receiver;
  return { connect, createReceiver, closeCalls };
}

describe("Substrate client cancellation via AbortSignal", () => {
  test("closes the socket and returns a cancelled result when the client disconnects mid-turn", async () => {
    const controller = new AbortController();
    const { connect, createReceiver, closeCalls } = makeFakeTransport(() => {
      controller.abort();
    });
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-cancel-1",
      makeRequest(),
      true,
      undefined,
      controller.signal,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(499);
    // Upstream work was actually stopped: the socket was closed for the abort.
    expect(
      closeCalls.some((call) => call.reason === "client disconnected"),
    ).toBeTrue();
  });

  test("does not open a socket when the request is already aborted before start", async () => {
    const controller = new AbortController();
    controller.abort();
    let connectCalls = 0;
    const connect: SubstrateWebSocketConnector = async () => {
      connectCalls += 1;
      return {} as unknown as FakeWebSocket;
    };
    const createReceiver: SubstrateReceiverFactory = () => {
      throw new Error("receiver must not be created for a pre-aborted turn");
    };
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-cancel-2",
      makeRequest(),
      true,
      undefined,
      controller.signal,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(499);
    expect(connectCalls).toBe(0);
  });
});

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
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
