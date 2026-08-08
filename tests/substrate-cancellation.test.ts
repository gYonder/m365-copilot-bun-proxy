import { describe, expect, test } from "bun:test";
import {
  CopilotSubstrateClient,
  buildInvocationPayload,
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
function makeFakeTransport(
  onBlockingRead: () => void,
  blockingReadCall = 2,
  throwOnCancel = false,
): {
  connect: SubstrateWebSocketConnector;
  createReceiver: SubstrateReceiverFactory;
  closeCalls: Array<{ code?: number; reason?: string }>;
  events: string[];
  sentFrames: string[];
  closedPromise: Promise<void>;
} {
  const closeCalls: Array<{ code?: number; reason?: string }> = [];
  const events: string[] = [];
  const sentFrames: string[] = [];
  let resolveClosed!: () => void;
  const closedPromise = new Promise<void>((resolve) => {
    resolveClosed = resolve;
  });
  let socketClosed = false;
  let resolvePendingRead: ((value: string | null) => void) | null = null;
  const ws = {
    readyState: 1, // WebSocket.OPEN
    send: (payload: string) => {
      if (throwOnCancel && payload.includes('"type":5')) {
        throw new Error("cancel send failed");
      }
      sentFrames.push(payload);
      events.push(`send:${payload}`);
    },
    close: (code?: number, reason?: string) => {
      if (!socketClosed) {
        socketClosed = true;
        (ws as { readyState: number }).readyState = 3; // WebSocket.CLOSED
        closeCalls.push({ code, reason });
        events.push("close");
        resolvePendingRead?.(null);
        resolvePendingRead = null;
        resolveClosed();
      }
    },
  };

  let nextCall = 0;
  const receiver = {
    next: async (_timeoutMs: number): Promise<string | null> => {
      nextCall += 1;
      if (socketClosed) {
        return null;
      }
      if (nextCall === blockingReadCall) {
        // The turn loop is now awaiting upstream output. Simulate the client
        // going away; the abort handler must close the socket.
        return new Promise<string | null>((resolve) => {
          resolvePendingRead = resolve;
          onBlockingRead();
        });
      }
      if (nextCall === 1) {
        return "{}"; // handshake frame, no error
      }
      // Once the abort closed the socket, the read completes with null and the
      // loop terminates. Otherwise it would block forever (the real "in-flight"
      // state) — which is exactly what cancellation must interrupt.
      return socketClosed ? null : await new Promise<string | null>(() => {});
    },
    dispose: () => {
      resolvePendingRead?.(null);
      resolvePendingRead = null;
    },
  };

  const connect: SubstrateWebSocketConnector = async () =>
    ws as unknown as FakeWebSocket;
  const createReceiver: SubstrateReceiverFactory = () => receiver;
  return {
    connect,
    createReceiver,
    closeCalls,
    events,
    sentFrames,
    closedPromise,
  };
}

describe("Substrate client cancellation via AbortSignal", () => {
  test("uses disconnectBehavior stop in the stream invocation payload", () => {
    const payload = buildInvocationPayload(
      makeRequest(),
      "conv-payload",
      "session-payload",
      "request-payload",
      true,
      createOptions(),
    );
    const argument = payload.arguments[0] as {
      disconnectBehavior?: string;
    };

    expect(argument.disconnectBehavior).toBe("stop");
    expect(JSON.stringify(payload)).not.toContain("continue");
  });

  test("closes the socket and returns a cancelled result when the client disconnects mid-turn", async () => {
    const controller = new AbortController();
    let abortAt = 0;
    const {
      connect,
      createReceiver,
      closeCalls,
      events,
      sentFrames,
      closedPromise,
    } = makeFakeTransport(() => {
      abortAt = Date.now();
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
    await closedPromise;
    expect(Date.now() - abortAt).toBeLessThan(1_000);
    // Upstream work was actually stopped: the socket was closed for the abort.
    expect(
      closeCalls.some((call) => call.reason === "client disconnected"),
    ).toBeTrue();
    const invocationFrame = sentFrames.find((frame) =>
      frame.includes('"target":"chat"'),
    );
    const invocationId = JSON.parse(
      invocationFrame?.replace(/\u001e$/, "") ?? "{}",
    ).invocationId;
    const stopFrame = sentFrames.find((frame) =>
      frame.includes('"type":5'),
    );
    const stopIndex = events.findIndex((event) =>
      event.includes('"type":5'),
    );
    expect(stopIndex).toBeGreaterThanOrEqual(0);
    expect(stopIndex).toBeLessThan(events.indexOf("close"));
    expect(sentFrames.filter((frame) => frame.includes('"type":5'))).toHaveLength(
      1,
    );
    expect(stopFrame).toBe(
      `${JSON.stringify({ type: 5, invocationId })}\u001e`,
    );
  });

  test("closes during handshake without sending a cancellation frame", async () => {
    const controller = new AbortController();
    const transport = makeFakeTransport(
      () => controller.abort(),
      1,
    );
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      transport.connect,
      transport.createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-handshake-cancel",
      makeRequest(),
      true,
      undefined,
      controller.signal,
    );

    expect(result.statusCode).toBe(499);
    await transport.closedPromise;
    expect(
      transport.sentFrames.some((frame) => frame.includes('"type":5')),
    ).toBeFalse();
    expect(transport.closeCalls).toHaveLength(1);
  });

  test("sends one cancellation frame and closes once on a double abort", async () => {
    const controller = new AbortController();
    const transport = makeFakeTransport(() => {
      controller.abort();
      controller.signal.dispatchEvent(new Event("abort"));
    });
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      transport.connect,
      transport.createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-double-cancel",
      makeRequest(),
      true,
      undefined,
      controller.signal,
    );

    expect(result.statusCode).toBe(499);
    await transport.closedPromise;
    expect(
      transport.sentFrames.filter((frame) => frame.includes('"type":5')),
    ).toHaveLength(1);
    expect(transport.closeCalls).toHaveLength(1);
  });

  test("closes when cancellation frame sending throws without an unhandled rejection", async () => {
    const controller = new AbortController();
    const transport = makeFakeTransport(
      () => controller.abort(),
      2,
      true,
    );
    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown) => {
      unhandled.push(reason);
    };
    process.on("unhandledRejection", onUnhandled);
    try {
      const client = new CopilotSubstrateClient(
        createOptions(),
        stubLogger,
        undefined,
        transport.connect,
        transport.createReceiver,
      );
      const result = await client.chat(
        makeJwtAuthHeader(),
        "conv-send-failure",
        makeRequest(),
        true,
        undefined,
        controller.signal,
      );

      expect(result.statusCode).toBe(499);
      await transport.closedPromise;
      expect(transport.closeCalls).toHaveLength(1);
      await Bun.sleep(0);
      expect(unhandled).toHaveLength(0);
    } finally {
      process.off("unhandledRejection", onUnhandled);
    }
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

describe("Substrate client lifecycle hardening", () => {
  test("echoes SignalR type-6 pings during a turn", async () => {
    const sentFrames: string[] = [];
    const { connect, createReceiver } = makeFrameTransport(
      [
        "{}",
        `${JSON.stringify({ type: 6 })}\u001e`,
        `${JSON.stringify({
          type: 1,
          target: "update",
          arguments: [
            {
              messages: [
                {
                  author: "bot",
                  messageType: "Chat",
                  messageId: "message-1",
                  text: "done",
                },
              ],
            },
          ],
        })}\u001e`,
        `${JSON.stringify({ type: 3, invocationId: "0" })}\u001e`,
      ],
      sentFrames,
    );
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-ping",
      makeRequest(),
      true,
    );

    expect(result.isSuccess).toBeTrue();
    expect(
      sentFrames.filter((frame) => frame.includes('"type":6')).length,
    ).toBe(2);
  });

  test("returns a typed failure for Disengaged messages", async () => {
    const { connect, createReceiver } = makeFrameTransport([
      "{}",
      `${JSON.stringify({
        type: 1,
        target: "update",
        arguments: [
          {
            messages: [
              {
                author: "bot",
                messageType: "Disengaged",
                messageId: "message-2",
                text: "Request declined.",
              },
            ],
          },
        ],
      })}\u001e`,
    ]);
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-disengaged",
      makeRequest(),
      true,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(502);
    expect(result.errorCode).toBe("substrate_disengaged");
    expect(result.rawBody).toContain("Request declined.");
  });

  test("folds repeated update frames without repeating output", async () => {
    const update = JSON.stringify({
      type: 1,
      target: "update",
      arguments: [{
        writeAtCursor: "once",
        messages: [{
          author: "bot",
          messageType: "Chat",
          messageId: "message-repeat",
          text: "once",
        }],
      }],
    }) + "\u001e";
    const { connect, createReceiver } = makeFrameTransport([
      "{}",
      update,
      update,
      JSON.stringify({ type: 3, invocationId: "0" }) + "\u001e",
    ]);
    const deltas: string[] = [];
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-repeat",
      makeRequest(),
      true,
      async (updateEvent) => {
        if (updateEvent.deltaText) deltas.push(updateEvent.deltaText);
      },
    );

    expect(result.isSuccess).toBeTrue();
    expect(result.assistantText).toBe("once");
    expect(deltas.join("")).toBe("once");
  });

  test("rejects partial output followed by a type-7 terminal error", async () => {
    const { connect, createReceiver } = makeFrameTransport([
      "{}",
      JSON.stringify({
        type: 1,
        target: "update",
        arguments: [{ writeAtCursor: "partial" }],
      }) + "\u001e",
      JSON.stringify({ type: 7, error: "terminal failure" }) + "\u001e",
    ]);
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-terminal-error",
      makeRequest(),
      true,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(502);
    expect(result.errorCode).toBe("substrate_terminal_error");
  });

  test("rejects socket closure without a successful terminal frame", async () => {
    const { connect, createReceiver } = makeFrameTransport([
      "{}",
      JSON.stringify({
        type: 1,
        target: "update",
        arguments: [{
          messages: [{
            author: "bot",
            messageType: "Chat",
            messageId: "message-incomplete",
            text: "not terminal",
          }],
        }],
      }) + "\u001e",
    ]);
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      createReceiver,
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-incomplete-terminal",
      makeRequest(),
      true,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.errorCode).toBe("substrate_incomplete_terminal");
  });

  test("honors an already-expired task-level deadline", async () => {
    let connectCalls = 0;
    const connect: SubstrateWebSocketConnector = async () => {
      connectCalls += 1;
      return {} as unknown as FakeWebSocket;
    };
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      connect,
      () => {
        throw new Error("receiver must not be created after task timeout");
      },
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-task-timeout",
      makeRequest(),
      true,
      undefined,
      null,
      Date.now() - 1,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(504);
    expect(connectCalls).toBe(0);
  });

  test("rechecks the task deadline after websocket connection", async () => {
    let receiverCalls = 0;
    const ws = {
      readyState: 1,
      send: () => {},
      close: () => {
        (ws as { readyState: number }).readyState = 3;
      },
    };
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      async () => {
        await Bun.sleep(5);
        return ws as unknown as FakeWebSocket;
      },
      () => {
        receiverCalls += 1;
        throw new Error("receiver must not be created after task timeout");
      },
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-connect-timeout",
      makeRequest(),
      true,
      undefined,
      null,
      Date.now() + 1,
    );

    expect(result.isSuccess).toBeFalse();
    expect(result.statusCode).toBe(504);
    expect(receiverCalls).toBe(0);
  });

  test("cancels while websocket connection is still pending", async () => {
    const controller = new AbortController();
    let resolveConnect!: (socket: FakeWebSocket) => void;
    let markConnectStarted!: () => void;
    let closeReason: string | undefined;
    const connectPromise = new Promise<FakeWebSocket>((resolve) => {
      resolveConnect = resolve;
    });
    const connectStarted = new Promise<void>((resolve) => {
      markConnectStarted = resolve;
    });
    const client = new CopilotSubstrateClient(
      createOptions(),
      stubLogger,
      undefined,
      async () => {
        markConnectStarted();
        return connectPromise;
      },
      () => {
        throw new Error("receiver must not be created after connect cancellation");
      },
    );

    const resultPromise = client.chat(
      makeJwtAuthHeader(),
      "conv-connect-cancel",
      makeRequest(),
      true,
      undefined,
      controller.signal,
    );
    await connectStarted;
    controller.abort();
    const result = await resultPromise;
    resolveConnect({
      readyState: 1,
      send: () => {},
      close: (_code?: number, reason?: string) => {
        closeReason = reason;
      },
    } as unknown as FakeWebSocket);
    await Bun.sleep(5);

    expect(result.statusCode).toBe(499);
    expect(closeReason).toBe("client disconnected");
  });

  test("reports a handshake deadline as 504 instead of socket closure", async () => {
    const options = createOptions();
    options.substrate.handshakeTimeoutSeconds = 0.001;
    const ws = {
      readyState: 1,
      send: () => {},
      close: () => {
        (ws as { readyState: number }).readyState = 3;
      },
    };
    const client = new CopilotSubstrateClient(
      options,
      stubLogger,
      undefined,
      async () => ws as unknown as FakeWebSocket,
      () => ({
        next: async () => {
          await Bun.sleep(5);
          return null;
        },
        dispose: () => {},
      }),
    );

    const result = await client.chat(
      makeJwtAuthHeader(),
      "conv-handshake-timeout",
      makeRequest(),
      true,
      undefined,
      null,
      Date.now() + 1_000,
    );

    expect(result.statusCode).toBe(504);
    expect(result.rawBody).toContain("handshake timed out");
  });
});

function makeFrameTransport(
  frames: string[],
  sentFrames: string[] = [],
): {
  connect: SubstrateWebSocketConnector;
  createReceiver: SubstrateReceiverFactory;
} {
  const ws = {
    readyState: 1,
    send: (payload: string) => {
      sentFrames.push(String(payload));
    },
    close: () => {
      (ws as { readyState: number }).readyState = 3;
    },
  };
  let index = 0;
  return {
    connect: async () => ws as unknown as FakeWebSocket,
    createReceiver: () => ({
      next: async () => frames[index++] ?? null,
      dispose: () => {},
    }),
  };
}

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
