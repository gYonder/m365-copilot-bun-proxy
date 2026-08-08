import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import { createProxyApp } from "../src/proxy/server";
import { ResponseStore } from "../src/proxy/response-store";
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
import { readSseEvents, tryParseJsonObject } from "../src/proxy/utils";

describe("Responses terminal invariant", () => {
  test("successful streaming Responses turn has one completed terminal and contiguous sequence numbers", async () => {
    const services = createServices({
      stream: async (onUpdate) => {
        await onUpdate({ deltaText: "hello", conversationId: null });
        return successResult("hello");
      },
    });

    const response = await postResponses(services);
    expect(response.status).toBe(200);
    const stream = await collectSse(response);
    const semanticEvents = stream.events.filter((event) =>
      event.type?.startsWith("response."),
    );
    const terminals = semanticEvents.filter((event) =>
      ["response.completed", "response.failed", "response.incomplete"].includes(
        String(event.type),
      ),
    );

    expect(terminals.map((event) => event.type)).toEqual(["response.completed"]);
    expect(semanticEvents.map((event) => event.sequence_number)).toEqual(
      semanticEvents.map((_, index) => index),
    );
    expect(stream.sawDone).toBeTrue();
    expect(stream.events.at(-1)?.type).toBe("response.completed");
  });

  test("upstream failure after partial emission produces response.failed without an error event", async () => {
    const services = createServices({
      stream: async (onUpdate) => {
        await onUpdate({ deltaText: "partial", conversationId: null });
        return failureResult("substrate_error");
      },
    });

    const stream = await collectSse(await postResponses(services));
    expect(stream.events.filter((event) => event.type === "error")).toHaveLength(0);
    expect(stream.events.filter((event) => event.type === "response.failed")).toHaveLength(1);
    expect(stream.events.filter((event) => event.type === "response.incomplete")).toHaveLength(0);
    expect(stream.events.at(-1)).toMatchObject({
      type: "response.failed",
      response: {
        status: "failed",
        error: { code: "transport_failed" },
      },
    });
    expect(stream.sawDone).toBeTrue();
  });

  test("upstream failure before any upstream emission still produces one terminal", async () => {
    const services = createServices({
      stream: async () => failureResult("substrate_error"),
    });

    const stream = await collectSse(await postResponses(services));
    const terminals = stream.events.filter((event) =>
      ["response.completed", "response.failed", "response.incomplete"].includes(
        String(event.type),
      ),
    );
    expect(terminals).toHaveLength(1);
    expect(terminals[0]?.type).toBe("response.failed");
    expect(stream.sawDone).toBeTrue();
  });

  test("missing upstream terminal frame produces response.incomplete with a reason", async () => {
    const services = createServices({
      options: { transport: TransportNames.Graph },
      graphStream: () =>
        new Response(
          new ReadableStream({
            start(controller) {
              controller.enqueue(
                new TextEncoder().encode(
                  `data: ${JSON.stringify({
                    id: "conv_graph",
                    messages: [{ text: "partial" }],
                  })}\n\n`,
                ),
              );
              controller.close();
            },
          }),
          { status: 200, headers: { "content-type": "text/event-stream" } },
        ),
    });

    const stream = await collectSse(
      await postResponses(services, TransportNames.Graph),
    );
    expect(stream.events.filter((event) => event.type === "response.incomplete")).toHaveLength(1);
    expect(stream.events.at(-1)).toMatchObject({
      type: "response.incomplete",
      response: {
        status: "incomplete",
        incomplete_details: { reason: "substrate_incomplete_terminal" },
      },
    });
    expect(stream.sawDone).toBeTrue();
  });

  test("Chat Completions streaming still terminates with [DONE]", async () => {
    const services = createServices({
      stream: async (onUpdate) => {
        await onUpdate({ deltaText: "hello", conversationId: null });
        return successResult("hello");
      },
    });

    const response = await services.app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "hello" }],
        }),
      }),
    );
    const stream = await collectSse(response);

    expect(stream.sawDone).toBeTrue();
    expect(stream.raw).toContain("[DONE]");
    expect(stream.events.some((event) => event.type === "response.completed")).toBeFalse();
  });
});

type StreamUpdate = {
  deltaText: string | null;
  conversationId: string | null;
};

type TestServices = ReturnType<typeof createServices>;

function createServices(config: {
  stream?: (
    onUpdate: (update: StreamUpdate) => Promise<void>,
  ) => Promise<ChatResult>;
  graphStream?: () => Response;
  options?: Partial<Pick<WrapperOptions, "transport">>;
}) {
  const options = createOptions();
  if (config.options?.transport) {
    options.transport = config.options.transport;
  }
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);
  const stream =
    config.stream ??
    (async () => successResult("hello"));

  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_graph",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => successResult("hello"),
    chatOverStream: async (): Promise<Response> =>
      config.graphStream?.() ?? new Response(null, { status: 500 }),
  } as unknown as CopilotGraphClient;

  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_substrate",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => successResult("hello"),
    chatStream: async (
      _authorizationHeader: string,
      _conversationId: string,
      _request: unknown,
      _createdConversation: boolean,
      onUpdate: (update: StreamUpdate) => Promise<void>,
    ): Promise<ChatResult> => stream(onUpdate),
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

  const app = createProxyApp({
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    tokenProvider,
  });
  return { app };
}

async function postResponses(
  services: TestServices,
  transport = TransportNames.Substrate,
): Promise<Response> {
  return services.app.fetch(
    new Request("http://localhost/v1/responses", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-m365-transport": transport,
      },
      body: JSON.stringify({
        model: "m365-copilot",
        stream: true,
        input: "hello",
      }),
    }),
  );
}

async function collectSse(response: Response): Promise<{
  events: JsonObject[];
  sawDone: boolean;
  raw: string;
}> {
  expect(response.body).not.toBeNull();
  const events: JsonObject[] = [];
  let sawDone = false;
  let raw = "";
  for await (const event of readSseEvents(response.body!)) {
    raw += `event: ${event.event}\ndata: ${event.data}\n\n`;
    if (event.data.trim().toLowerCase() === "[done]") {
      sawDone = true;
      continue;
    }
    const parsed = tryParseJsonObject(event.data);
    if (parsed) {
      events.push(parsed);
    }
  }
  return { events, sawDone, raw };
}

function successResult(text: string): ChatResult {
  return {
    isSuccess: true,
    statusCode: 200,
    responseJson: {
      id: "conv_substrate",
      messages: [{ text }],
    },
    rawBody: "{}",
    assistantText: text,
    conversationId: "conv_substrate",
  };
}

function failureResult(errorCode: string): ChatResult {
  return {
    isSuccess: false,
    statusCode: 502,
    responseJson: null,
    rawBody: JSON.stringify({ code: errorCode }),
    assistantText: null,
    conversationId: null,
    errorCode,
  };
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Mapped,
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
