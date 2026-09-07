import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  computeResponsesReplayIdentityKey,
  createProxyApp,
} from "../src/proxy/server";
import { tryParseResponsesRequest } from "../src/proxy/request-parser";
import { ResponseStore } from "../src/proxy/response-store";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import {
  LogLevels,
  OpenAiTransformModes,
  ToolChoiceModes,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type ParsedResponsesRequest,
  type ResponsesProtocolIdentity,
  type WrapperOptions,
} from "../src/proxy/types";
import {
  isJsonObject,
  readSseEvents,
  tryGetString,
  tryParseJsonObject,
} from "../src/proxy/utils";

const privateMarker = "\uE000cite\uE001turn1search1\uE002";

describe("final strict-boundary release review", () => {
  test.each([
    ["chat", false],
    ["chat", true],
    ["responses", false],
    ["responses", true],
  ] as const)(
    "uses the accepted Graph correction conversation for %s %s",
    async (endpoint, stream) => {
      const chatConversationIds: string[] = [];
      let createConversationCount = 0;
      const services = createGraphServices(
        (conversationId, _payload, callIndex) => {
          chatConversationIds.push(conversationId);
          return graphResult(
            callIndex === 1
              ? invalidPayload(endpoint)
              : validPayload(endpoint, "accepted"),
          );
        },
        () => {
          createConversationCount += 1;
          return `conv-${createConversationCount}`;
        },
      );
      const app = createProxyApp(services);
      const headers = {
        "content-type": "application/json",
        "x-m365-transport": TransportNames.Graph,
        "x-m365-conversation-key": "release-review",
      };
      const body =
        endpoint === "chat"
          ? {
              model: "m365-copilot",
              stream,
              messages: [{ role: "user", content: "Say accepted." }],
            }
          : {
              model: "m365-copilot",
              stream,
              input: "Say accepted.",
            };

      const response = await app.fetch(
        new Request(
          `http://localhost/v1/${endpoint === "chat" ? "chat/completions" : "responses"}`,
          {
            method: "POST",
            headers,
            body: JSON.stringify(body),
          },
        ),
      );

      expect(response.status).toBe(200);
      expect(chatConversationIds).toEqual(["conv-1", "conv-2"]);
      expect(response.headers.get("x-m365-conversation-id")).toBe("conv-2");

      const events = stream ? await collectSse(response) : [];
      if (endpoint === "chat") {
        if (stream) {
          expect(
            events
              .filter((event) => isJsonObject(event) && event.object === "chat.completion.chunk")
              .every((event) => event.conversation_id === "conv-2"),
          ).toBeTrue();
        } else {
          const responseBody = (await response.json()) as JsonObject;
          expect(responseBody.conversation_id).toBe("conv-2");
        }

        const followUp = await app.fetch(
          new Request("http://localhost/v1/chat/completions", {
            method: "POST",
            headers,
            body: JSON.stringify({
              model: "m365-copilot",
              stream: false,
              messages: [{ role: "user", content: "Follow up." }],
            }),
          }),
        );
        expect(followUp.status).toBe(200);
        expect(chatConversationIds[2]).toBe("conv-2");
        return;
      }

      const responseBody = stream
        ? findCompletedResponse(events)
        : ((await response.json()) as JsonObject);
      expect(responseBody.conversation).toBe("conv-2");
      expect(responseBody.conversation_id).toBe("conv-2");
      const responseId = tryGetString(responseBody, "id");
      expect(responseId).toBeTruthy();
      expect(services.responseStore.tryGetConversationLink(responseId!)).toBe(
        "conv-2",
      );
      if (stream) {
        expect(
          events
            .filter((event) => isJsonObject(event) && event.type === "response.in_progress")
            .every((event) => (event.response as JsonObject).conversation === "conv-2"),
        ).toBeTrue();
      }
    },
  );

  test("parses flat function and named custom Responses tool choices with exact kinds", () => {
    const functionParsed = parseResponsesWithTools({
      type: "function",
      name: "exec",
    }, [{
      type: "function",
      name: "exec",
      parameters: { type: "object" },
    }]);
    expect(functionParsed.base.tooling).toMatchObject({
      toolChoiceMode: ToolChoiceModes.Function,
      toolChoiceFunctionName: "exec",
      toolChoiceToolType: "function",
    });

    const customParsed = parseResponsesWithTools({
      type: "custom",
      custom: { name: "apply_patch" },
    }, [{
      type: "custom",
      name: "apply_patch",
      format: { type: "grammar", syntax: "lark" },
    }]);
    expect(customParsed.base.tooling).toMatchObject({
      toolChoiceMode: ToolChoiceModes.Function,
      toolChoiceFunctionName: "apply_patch",
      toolChoiceToolType: "custom",
    });

    const chatShapeParsed = parseResponsesWithTools({
      type: "function",
      function: { name: "exec" },
    }, [{
      type: "function",
      name: "exec",
      parameters: { type: "object" },
    }]);
    expect(chatShapeParsed.base.tooling.toolChoiceFunctionName).toBe("exec");
    expect(chatShapeParsed.base.tooling.toolChoiceToolType).toBe("function");
  });

  test("corrects a forced function choice instead of accepting another function", async () => {
    const chatConversationIds: string[] = [];
    const services = createGraphServices(
      (conversationId, _payload, callIndex) => {
        chatConversationIds.push(conversationId);
        return graphResult(
          callIndex === 1
            ? validFunctionCallPayload("other", "call_wrong")
            : validFunctionCallPayload("exec", "call_exec"),
        );
      },
      (index) => `conv-${index}`,
    );
    const response = await createProxyApp(services).fetch(
      responsesRequest({
        input: "Use exec.",
        tools: [
          { type: "function", name: "exec", parameters: { type: "object" } },
          { type: "function", name: "other", parameters: { type: "object" } },
        ],
        tool_choice: { type: "function", name: "exec" },
      }),
    );

    expect(response.status).toBe(200);
    expect(chatConversationIds).toEqual(["conv-1", "conv-2"]);
    expect(response.headers.get("x-m365-conversation-id")).toBe("conv-2");
    const body = (await response.json()) as JsonObject;
    expect((body.output as JsonObject[])[0]).toMatchObject({
      type: "function_call",
      name: "exec",
    });
  });

  test("rejects a named custom choice with a function-kind call and corrects it", async () => {
    const expectedInput =
      "*** Begin Patch\n*** Add File: exact.txt\n+exact\n*** End Patch";
    let callCount = 0;
    const services = createGraphServices(
      (conversationId) => {
        callCount += 1;
        return graphResult(
          callCount === 1
            ? validFunctionCallPayload("apply_patch", "call_wrong")
            : validCustomCallPayload("apply_patch", expectedInput),
        );
      },
      (index) => `conv-${index}`,
    );
    const response = await createProxyApp(services).fetch(
      responsesRequest({
        input: "Apply the patch.",
        tools: [{
          type: "custom",
          name: "apply_patch",
          format: { type: "grammar", syntax: "lark" },
        }],
        tool_choice: { type: "custom", name: "apply_patch" },
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    expect((body.output as JsonObject[])[0]).toMatchObject({
      type: "custom_tool_call",
      name: "apply_patch",
      input: expectedInput,
    });
  });

  test("does not commit a cancelled valid Responses turn and retries upstream later", async () => {
    const controller = new AbortController();
    let callCount = 0;
    const services = createGraphServices(
      (conversationId) => {
        callCount += 1;
        if (callCount === 1) {
          controller.abort();
        }
        return graphResult(validFunctionCallPayload("exec", "call_once"));
      },
      (index) => `conv-${index}`,
    );
    const requestBody = {
      model: "m365-copilot",
      stream: true,
      conversation: "conv-fixed",
      client_metadata: {
        thread_id: "thread-cancel",
        session_id: "session-cancel",
        turn_id: "turn-cancel",
      },
      input: "Run exec.",
      tools: [{
        type: "function",
        name: "exec",
        parameters: { type: "object" },
      }],
      tool_choice: "required",
    };
    const first = await createProxyApp(services).fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        signal: controller.signal,
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );
    expect(first.status).toBe(200);
    expect(await collectSse(first)).toEqual([]);

    const parsed = tryParseResponsesRequest(requestBody, services.options);
    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    const replayKey = computeResponsesReplayIdentityKey(
      parsed.request.protocolIdentity,
      requestBody,
      TransportNames.Graph,
      "unused",
    );
    expect(services.responseStore.tryGetProtocolReplay(replayKey)).toBeNull();

    const second = await createProxyApp(services).fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({ ...requestBody, stream: false }),
      }),
    );
    expect(second.status).toBe(200);
    expect(callCount).toBe(2);
    const secondBody = (await second.json()) as JsonObject;
    expect((secondBody.output as JsonObject[])[0]?.type).toBe("function_call");
  });

  test("sanitizes only Responses output text before nonstream and stream projection", async () => {
    const outputText = `Before ${privateMarker} after {\"code\":\"kept\"}`;
    const functionArguments = `{\"cmd\":\"echo ${privateMarker}\"}`;
    const customInput =
      `*** Begin Patch\n*** Add File: marker.txt\n+${privateMarker}\n*** End Patch`;
    const providerResponse = {
      object: "response",
      status: "completed",
      output: [
        {
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: outputText }],
        },
        {
          type: "function_call",
          status: "completed",
          call_id: "call_exec",
          name: "exec",
          arguments: functionArguments,
        },
        {
          type: "custom_tool_call",
          status: "completed",
          call_id: "call_patch",
          name: "apply_patch",
          input: customInput,
        },
      ],
      output_text: outputText,
    };
    let conversationCount = 0;
    const services = createGraphServices(
      () => graphResult(providerResponse),
      () => `conv-${++conversationCount}`,
    );
    const app = createProxyApp(services);
    const requestBody = {
      model: "m365-copilot",
      input: "Return the mixed response.",
      tools: [
        { type: "function", name: "exec", parameters: { type: "object" } },
        { type: "custom", name: "apply_patch", format: { type: "grammar" } },
      ],
      parallel_tool_calls: true,
    };

    const nonstream = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({ ...requestBody, stream: false }),
      }),
    );
    const nonstreamBody = (await nonstream.json()) as JsonObject;
    const sanitizedText = "Before after {\"code\":\"kept\"}";
    expect(nonstreamBody.output_text).toBe(sanitizedText);
    expect(
      ((nonstreamBody.output as JsonObject[])[0]?.content as JsonObject[])[0]?.text,
    ).toBe(sanitizedText);
    expect((nonstreamBody.output as JsonObject[])[1]?.arguments).toBe(
      functionArguments,
    );
    expect((nonstreamBody.output as JsonObject[])[2]?.input).toBe(customInput);

    const streamResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
          "x-m365-new-conversation": "true",
        },
        body: JSON.stringify({ ...requestBody, stream: true }),
      }),
    );
    const events = await collectSse(streamResponse);
    const deltas = events
      .filter((event) => event.type === "response.output_text.delta")
      .map((event) => String(event.delta ?? ""))
      .join("");
    const textDone = events.find(
      (event) => event.type === "response.output_text.done",
    );
    const completed = findCompletedResponse(events);
    expect(deltas).toBe(sanitizedText);
    expect(textDone?.text).toBe(sanitizedText);
    expect(
      (((completed.output as JsonObject[])[0]?.content as JsonObject[])[0]?.text),
    ).toBe(sanitizedText);
    expect((completed.output as JsonObject[])[1]?.arguments).toBe(
      functionArguments,
    );
    expect((completed.output as JsonObject[])[2]?.input).toBe(customInput);
  });
});

function createGraphServices(
  onChat: (
    conversationId: string,
    payload: JsonObject,
    callIndex: number,
  ) => ChatResult,
  createConversationId: (index: number) => string,
): Parameters<typeof createProxyApp>[0] {
  const options = createOptions();
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);
  let chatCallIndex = 0;
  let createCallIndex = 0;
  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: createConversationId(++createCallIndex),
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
      payload: JsonObject,
    ): Promise<ChatResult> =>
      onChat(conversationId, payload, ++chatCallIndex),
    chatOverStream: async (): Promise<Response> =>
      new Response(null, { status: 500 }),
  } as unknown as CopilotGraphClient;
  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "substrate-unused",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("Substrate chat is not used by this test.");
    },
    chatStream: async (): Promise<ChatResult> => {
      throw new Error("Substrate streaming is not used by this test.");
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
  return {
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    tokenProvider,
  };
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
    createConversationPath: "/beta/copilot/conversations",
    chatPathTemplate: "/beta/copilot/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate:
      "/beta/copilot/conversations/{conversationId}/chatOverStream",
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
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
    logStdout: false,
    confabRetries: 0,
    msalAuth: false,
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

function graphResult(payload: JsonObject): ChatResult {
  const assistantText = toMarkdownJson(payload);
  return {
    isSuccess: true,
    statusCode: 200,
    responseJson: {
      id: "graph-response",
      messages: [{ text: "prompt" }, { text: assistantText }],
    },
    rawBody: "{}",
    assistantText: null,
    conversationId: null,
  };
}

function invalidPayload(endpoint: "chat" | "responses"): JsonObject {
  return endpoint === "chat"
    ? { object: "not-chat" }
    : { object: "not-response" };
}

function validPayload(
  endpoint: "chat" | "responses",
  text: string,
): JsonObject {
  return endpoint === "chat"
    ? {
        object: "chat.completion",
        choices: [{
          index: 0,
          finish_reason: "stop",
          message: { role: "assistant", content: text },
        }],
      }
    : {
        object: "response",
        status: "completed",
        output: [{
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text }],
        }],
        output_text: text,
      };
}

function validFunctionCallPayload(name: string, callId: string): JsonObject {
  return {
    object: "response",
    status: "completed",
    output: [{
      type: "function_call",
      status: "completed",
      call_id: callId,
      name,
      arguments: "{}",
    }],
  };
}

function validCustomCallPayload(name: string, input: string): JsonObject {
  return {
    object: "response",
    status: "completed",
    output: [{
      type: "custom_tool_call",
      status: "completed",
      call_id: "call_custom",
      name,
      input,
    }],
  };
}

function parseResponsesWithTools(
  toolChoice: JsonObject,
  tools: JsonValueLike[],
): ParsedResponsesRequest {
  const parsed = tryParseResponsesRequest({
    model: "m365-copilot",
    input: "Use a tool.",
    tools,
    tool_choice: toolChoice,
  }, createOptions());
  expect(parsed.ok).toBeTrue();
  if (!parsed.ok) {
    throw new Error(parsed.error);
  }
  return parsed.request;
}

type JsonValueLike = JsonObject;

function responsesRequest(body: JsonObject): Request {
  return new Request("http://localhost/v1/responses", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-m365-transport": TransportNames.Graph,
    },
    body: JSON.stringify({ model: "m365-copilot", stream: false, ...body }),
  });
}

async function collectSse(response: Response): Promise<JsonObject[]> {
  const events: JsonObject[] = [];
  if (!response.body) {
    return events;
  }
  for await (const event of readSseEvents(response.body)) {
    const data = event.data.trim();
    if (!data || data.toLowerCase() === "[done]") {
      continue;
    }
    const parsed = tryParseJsonObject(data);
    if (parsed) {
      events.push(parsed);
    }
  }
  return events;
}

function findCompletedResponse(events: JsonObject[]): JsonObject {
  const completed = events.find((event) => event.type === "response.completed");
  if (!completed || !isJsonObject(completed.response)) {
    throw new Error("Missing response.completed event.");
  }
  return completed.response;
}

function toMarkdownJson(payload: JsonObject): string {
  return `\`\`\`json\n${JSON.stringify(payload, null, 2)}\n\`\`\``;
}
