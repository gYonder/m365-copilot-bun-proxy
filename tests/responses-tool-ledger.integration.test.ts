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

const taskMessage = {
  type: "message",
  role: "user",
  content: [{ type: "input_text", text: "Complete the coding task." }],
};

const tools = [
  {
    type: "function",
    name: "read_file",
    description: "Read a file.",
    parameters: {
      type: "object",
      properties: { path: { type: "string" } },
      required: ["path"],
      additionalProperties: false,
    },
  },
];

describe("Responses tool ledger integration", () => {
  test("supports more than eight rounds and resets separated repetitions", async () => {
    let upstreamCalls = 0;
    const app = createApp(() => {
      upstreamCalls += 1;
      const path = upstreamCalls % 2 === 0 ? "other.txt" : "same.txt";
      return toolResponse(`call_${upstreamCalls}`, { path });
    });

    let previousResponseId: string | null = null;
    let previousCallId: string | null = null;
    for (let round = 1; round <= 10; round += 1) {
      const input: JsonObject[] = [taskMessage];
      if (previousCallId) {
        input.push(
          {
            type: "function_call",
            call_id: previousCallId,
            name: "read_file",
            arguments: JSON.stringify({ path: round % 2 === 0 ? "same.txt" : "other.txt" }),
          },
          {
            type: "function_call_output",
            call_id: previousCallId,
            output: "ok",
          },
        );
      }
      const response = await post(app, {
        model: "m365-copilot",
        input,
        previous_response_id: previousResponseId,
        tools,
        tool_choice: "required",
      });
      expect(response.status).toBe(200);
      const body = (await response.json()) as JsonObject;
      expect(body.status).toBe("completed");
      const call = firstFunctionCall(body);
      expect(call).not.toBeNull();
      previousCallId = call?.call_id as string;
      previousResponseId = body.id as string;
    }

    expect(upstreamCalls).toBe(10);
  });

  test("fails after five consecutive identical calls with one terminal response", async () => {
    let upstreamCalls = 0;
    const app = createApp(() => {
      upstreamCalls += 1;
      return toolResponse(`call_${upstreamCalls}`, { path: "same.txt" });
    });

    let previousResponseId: string | null = null;
    let previousCallId: string | null = null;
    let failureBody: JsonObject | null = null;
    for (let round = 1; round <= 5; round += 1) {
      const input: JsonObject[] = [taskMessage];
      if (previousCallId) {
        input.push(
          {
            type: "function_call",
            call_id: previousCallId,
            name: "read_file",
            arguments: JSON.stringify({ path: "same.txt" }),
          },
          {
            type: "function_call_output",
            call_id: previousCallId,
            output: "ok",
          },
        );
      }
      const response = await post(app, {
        model: "m365-copilot",
        input,
        previous_response_id: previousResponseId,
        tools,
        tool_choice: "required",
      });
      expect(response.status).toBe(200);
      const body = (await response.json()) as JsonObject;
      if (body.status === "failed") {
        failureBody = body;
        break;
      }
      previousCallId = firstFunctionCall(body)?.call_id as string;
      previousResponseId = body.id as string;
    }

    expect(upstreamCalls).toBe(5);
    expect(failureBody).toMatchObject({
      status: "failed",
      error: { code: "fallback_exhausted" },
    });
  });

  test("rejects a duplicate completed result with one semantic terminal", async () => {
    let upstreamCalls = 0;
    const app = createApp(() => {
      upstreamCalls += 1;
      return upstreamCalls === 1
        ? toolResponse("call_1", { path: "same.txt" })
        : finalResponse("done");
    });

    const first = await post(app, {
      model: "m365-copilot",
      input: [taskMessage],
      tools,
      tool_choice: "required",
    });
    const firstBody = (await first.json()) as JsonObject;
    const continuation = {
      model: "m365-copilot",
      previous_response_id: firstBody.id,
      client_metadata: { turn_id: "turn-2" },
      input: [
        taskMessage,
        {
          type: "function_call",
          call_id: "call_1",
          name: "read_file",
          arguments: JSON.stringify({ path: "same.txt" }),
        },
        { type: "function_call_output", call_id: "call_1", output: "ok" },
      ],
      tools,
      tool_choice: "auto",
    };
    const second = await post(app, continuation);
    expect(second.status).toBe(200);
    expect((await second.json()).status).toBe("completed");

    const duplicate = await post(
      app,
      { ...continuation, client_metadata: { turn_id: "turn-3" }, stream: true },
    );
    const events = await collectEvents(duplicate);
    const terminals = events.filter((event) =>
      ["response.completed", "response.failed", "response.incomplete"].includes(
        String(event.type),
      ),
    );
    expect(terminals).toHaveLength(1);
    expect(terminals[0]).toMatchObject({
      type: "response.failed",
      response: { error: { code: "duplicate_suppressed" } },
    });
    expect(upstreamCalls).toBe(2);
  });

  test("accepts an unknown result with a cold ledger after restart", async () => {
    const app = createApp(() => finalResponse("accepted"));
    const response = await post(
      app,
      {
        model: "m365-copilot",
        input: [
          taskMessage,
          {
            type: "function_call_output",
            call_id: "call_from_before_restart",
            output: "large but legitimate result",
          },
        ],
      },
      { "x-m365-conversation-id": "conv_cold" },
    );

    expect(response.status).toBe(200);
    expect((await response.json()).status).toBe("completed");
  });

  test("accepts parallel results returned in reverse order", async () => {
    let upstreamCalls = 0;
    const app = createApp(() => {
      upstreamCalls += 1;
      return upstreamCalls === 1
        ? {
            ...toolResponse("call_1", { path: "one.txt" }),
            output: [
              {
                type: "function_call",
                call_id: "call_1",
                name: "read_file",
                arguments: JSON.stringify({ path: "one.txt" }),
              },
              {
                type: "function_call",
                call_id: "call_2",
                name: "read_file",
                arguments: JSON.stringify({ path: "two.txt" }),
              },
            ],
          }
        : finalResponse("done");
    });

    const first = await post(app, {
      model: "m365-copilot",
      input: [taskMessage],
      tools,
      tool_choice: "required",
    });
    const firstBody = (await first.json()) as JsonObject;
    const second = await post(app, {
      model: "m365-copilot",
      previous_response_id: firstBody.id,
      input: [
        taskMessage,
        {
          type: "function_call",
          call_id: "call_1",
          name: "read_file",
          arguments: JSON.stringify({ path: "one.txt" }),
        },
        {
          type: "function_call",
          call_id: "call_2",
          name: "read_file",
          arguments: JSON.stringify({ path: "two.txt" }),
        },
        { type: "function_call_output", call_id: "call_2", output: "two" },
        { type: "function_call_output", call_id: "call_1", output: "one" },
      ],
      tools,
      tool_choice: "auto",
    });

    expect(second.status).toBe(200);
    expect((await second.json()).status).toBe("completed");
    expect(upstreamCalls).toBe(2);
  });
});

function createApp(
  onChat: () => JsonObject,
): ReturnType<typeof createProxyApp> {
  const options = createOptions();
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);
  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_ledger",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => ({
      isSuccess: true,
      statusCode: 200,
      responseJson: null,
      rawBody: "{}",
      assistantText: `\`\`\`json\n${JSON.stringify(onChat())}\n\`\`\``,
      conversationId: "conv_ledger",
    }),
    chatOverStream: async (): Promise<Response> =>
      new Response(null, { status: 500 }),
  } as unknown as CopilotGraphClient;
  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_unused",
      rawBody: "{}",
    }),
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
  return app;
}

async function post(
  app: ReturnType<typeof createApp>,
  body: JsonObject,
  extraHeaders: Record<string, string> = {},
): Promise<Response> {
  return app.fetch(
    new Request("http://localhost/v1/responses", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-m365-transport": TransportNames.Graph,
        ...extraHeaders,
      },
      body: JSON.stringify(body),
    }),
  );
}

function firstFunctionCall(body: JsonObject): JsonObject | null {
  if (!Array.isArray(body.output)) return null;
  return (
    body.output.find(
      (item): item is JsonObject =>
        typeof item === "object" &&
        item !== null &&
        !Array.isArray(item) &&
        ((item as JsonObject).type === "function_call" ||
          (item as JsonObject).type === "custom_tool_call"),
    ) ?? null
  );
}

function toolResponse(callId: string, args: JsonObject): JsonObject {
  return {
    id: `upstream_${callId}`,
    object: "response",
    output: [
      {
        type: "function_call",
        call_id: callId,
        name: "read_file",
        arguments: JSON.stringify(args),
      },
    ],
  };
}

function finalResponse(text: string): JsonObject {
  return {
    id: "upstream_final",
    object: "response",
    output: [
      {
        type: "message",
        role: "assistant",
        content: [{ type: "output_text", text }],
      },
    ],
  };
}

async function collectEvents(response: Response): Promise<JsonObject[]> {
  const events: JsonObject[] = [];
  for await (const event of readSseEvents(response.body!)) {
    if (event.data.trim().toLowerCase() === "[done]") continue;
    const parsed = tryParseJsonObject(event.data);
    if (parsed) events.push(parsed);
  }
  return events;
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Error,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Graph,
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
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}
