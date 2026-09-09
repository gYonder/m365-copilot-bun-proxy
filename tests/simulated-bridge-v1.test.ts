import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  appendSimulatedProtocolCorrection,
  buildSimulatedRegenerationPrompt,
  tryParseOpenAiRequest,
  tryParseResponsesRequest,
} from "../src/proxy/request-parser";
import { ResponseStore } from "../src/proxy/response-store";
import { createProxyApp } from "../src/proxy/server";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import {
  LogLevels,
  OpenAiTransformModes,
  ResponseFormatTypes,
  SimulatedOutputProtocols,
  ToolChoiceModes,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type OpenAiToolDefinition,
  type SimulatedOutputProtocol,
  type WrapperOptions,
} from "../src/proxy/types";
import {
  isJsonObject,
  readSseEvents,
  tryGetString,
  tryParseJsonObject,
} from "../src/proxy/utils";

describe("simulated action output protocol integration (bridge_v1)", () => {
  test("default legacy prompt and behavior remain unchanged", async () => {
    const defaultOptions = createOptions();
    expect(defaultOptions.simulatedOutputProtocol).toBe(
      SimulatedOutputProtocols.Legacy,
    );

    const parsedResponses = tryParseResponsesRequest(
      { model: "m365-copilot", input: "Summarize this task." },
      defaultOptions,
    );
    expect(parsedResponses.ok).toBeTrue();
    if (!parsedResponses.ok) return;
    const prompt = parsedResponses.request.base.promptText;
    expect(prompt).not.toContain("M365_FINAL_V1");
    expect(prompt).not.toContain("M365_FUNCTION_TOOL_CALL_V1");
    expect(prompt).not.toContain("M365_CUSTOM_TOOL_CALL_V1");
    expect(prompt).toContain("Return the complete assistant answer as plain text.");

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildChatResult(conversationId, payload, "Legacy plain text response");
      }),
    );
    const response = await postResponses(app, {
      model: "m365-copilot",
      stream: false,
      input: "Hello legacy",
    });
    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.status).toBe("completed");
    expect(body.output_text).toBe("Legacy plain text response");
    expect(callCount).toBe(1);
  });

  test("bridge_v1 prompt advertises exact byte-zero forms for responses only", () => {
    const v1Options = createOptions(SimulatedOutputProtocols.BridgeV1);
    const parsedResponses = tryParseResponsesRequest(
      {
        model: "m365-copilot",
        input: "Run file action",
        tools: [functionTool("read_file"), customTool("apply_diff")],
      },
      v1Options,
    );
    expect(parsedResponses.ok).toBeTrue();
    if (!parsedResponses.ok) return;
    const responsesPrompt = parsedResponses.request.base.promptText;
    expect(responsesPrompt).toContain("M365_FINAL_V1\\n<raw final text>");
    expect(responsesPrompt).toContain(
      "M365_FUNCTION_TOOL_CALL_V1\\n<exact tool name>\\n<one JSON argument object>",
    );
    expect(responsesPrompt).toContain(
      "M365_CUSTOM_TOOL_CALL_V1\\n<exact tool name>\\n<raw input to EOF>",
    );
    expect(responsesPrompt).toContain(
      "V1 represents only final text or exactly one function/custom call. If multiple/parallel calls, refusal, failed/incomplete response, mixed output, or JSON response format are required, model must use the existing strict legacy endpoint JSON envelope.",
    );
    expect(responsesPrompt).toContain(
      "Existing strict legacy endpoint JSON remains compatibility fallback under bridge_v1.",
    );

    const parsedChat = tryParseOpenAiRequest(
      {
        model: "m365-copilot",
        messages: [{ role: "user", content: "Hello chat" }],
        tools: [functionTool("read_file")],
      },
      v1Options,
    );
    expect(parsedChat.ok).toBeTrue();
    if (!parsedChat.ok) return;
    const chatPrompt = parsedChat.request.promptText;
    expect(chatPrompt).not.toContain("M365_FINAL_V1");
    expect(chatPrompt).not.toContain("M365_FUNCTION_TOOL_CALL_V1");
    expect(chatPrompt).not.toContain("M365_CUSTOM_TOOL_CALL_V1");
  });

  test("final text, function call, and custom tool call work through /v1/responses", async () => {
    const finalText = "Preserved final message with\nexact trailing whitespace   \n";
    const appFinal = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(conversationId, payload, `M365_FINAL_V1\n${finalText}`),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const resFinal = await postResponses(appFinal, {
      model: "m365-copilot",
      stream: false,
      input: "Say hello",
    });
    expect(resFinal.status).toBe(200);
    const bodyFinal = (await resFinal.json()) as JsonObject;
    expect(bodyFinal.status).toBe("completed");
    expect(bodyFinal.output_text).toBe(finalText);
    const outputFinal = bodyFinal.output as JsonObject[];
    expect(outputFinal).toHaveLength(1);
    expect(outputFinal[0]?.type).toBe("message");

    const appFunction = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(
            conversationId,
            payload,
            'M365_FUNCTION_TOOL_CALL_V1\nexec\n{"command":"git status"}',
          ),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const resFunction = await postResponses(appFunction, {
      model: "m365-copilot",
      stream: false,
      input: "Check git status",
      tools: [functionTool("exec")],
    });
    expect(resFunction.status).toBe(200);
    const bodyFunction = (await resFunction.json()) as JsonObject;
    expect(bodyFunction.status).toBe("completed");
    const outputFunction = bodyFunction.output as JsonObject[];
    expect(outputFunction).toHaveLength(1);
    expect(outputFunction[0]?.type).toBe("function_call");
    expect(outputFunction[0]?.name).toBe("exec");
    expect(outputFunction[0]?.arguments).toBe('{"command":"git status"}');

    const customInput = "<<<<<<< SEARCH\nfoo\n=======\nbar\n>>>>>>> REPLACE\n";
    const appCustom = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(
            conversationId,
            payload,
            `M365_CUSTOM_TOOL_CALL_V1\napply_diff\n${customInput}`,
          ),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const resCustom = await postResponses(appCustom, {
      model: "m365-copilot",
      stream: false,
      input: "Apply patch",
      tools: [customTool("apply_diff")],
    });
    expect(resCustom.status).toBe(200);
    const bodyCustom = (await resCustom.json()) as JsonObject;
    expect(bodyCustom.status).toBe("completed");
    const outputCustom = bodyCustom.output as JsonObject[];
    expect(outputCustom).toHaveLength(1);
    expect(outputCustom[0]?.type).toBe("custom_tool_call");
    expect(outputCustom[0]?.name).toBe("apply_diff");
    expect(outputCustom[0]?.input).toBe(customInput);

    const arbitraryBytes = "raw payload byte zero \n multi-line   \t \r\n";
    const appArbitrary = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(
            conversationId,
            payload,
            `M365_CUSTOM_TOOL_CALL_V1\ncustom_runner\n${arbitraryBytes}`,
          ),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const resArbitrary = await postResponses(appArbitrary, {
      model: "m365-copilot",
      stream: false,
      input: "Run arbitrary tool",
      tools: [customTool("custom_runner")],
    });
    expect(resArbitrary.status).toBe(200);
    const bodyArbitrary = (await resArbitrary.json()) as JsonObject;
    expect(bodyArbitrary.status).toBe("completed");
    const outputArbitrary = bodyArbitrary.output as JsonObject[];
    expect(outputArbitrary).toHaveLength(1);
    expect(outputArbitrary[0]?.type).toBe("custom_tool_call");
    expect(outputArbitrary[0]?.name).toBe("custom_runner");
    expect(outputArbitrary[0]?.input).toBe(arbitraryBytes);
  });

  test("streaming /v1/responses works with bridge_v1 V1 frame", async () => {
    const app = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(
            conversationId,
            payload,
            'M365_FUNCTION_TOOL_CALL_V1\nexec\n{"command":"git diff"}',
          ),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const res = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: "Run git diff",
          tools: [functionTool("exec")],
        }),
      }),
    );
    expect(res.status).toBe(200);
    let completed: JsonObject | null = null;
    for await (const event of readSseEvents(res.body!)) {
      const data = event.data.trim();
      if (!data || data.toLowerCase() === "[done]") continue;
      const parsed = tryParseJsonObject(data);
      if (parsed && tryGetString(parsed, "type") === "response.completed") {
        completed = isJsonObject(parsed.response) ? (parsed.response as JsonObject) : null;
      }
    }
    expect(completed?.status).toBe("completed");
    const output = completed?.output as JsonObject[];
    expect(output).toHaveLength(1);
    expect(output[0]?.type).toBe("function_call");
    expect(output[0]?.name).toBe("exec");
    expect(output[0]?.arguments).toBe('{"command":"git diff"}');
  });

  test("strict legacy JSON fallback still works under bridge_v1", async () => {
    const fallbackPayload = JSON.stringify({
      object: "response",
      status: "completed",
      output: [
        {
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: "Legacy JSON fallback text" }],
        },
      ],
    });
    const app = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(conversationId, payload, `\`\`\`json\n${fallbackPayload}\n\`\`\``),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );

    const res = await postResponses(app, {
      model: "m365-copilot",
      stream: false,
      input: "Provide fallback response",
    });
    expect(res.status).toBe(200);
    const body = (await res.json()) as JsonObject;
    expect(body.status).toBe("completed");
    expect(body.output_text).toBe("Legacy JSON fallback text");
  });

  test("malformed recognized V1 gets one correction and is never treated as raw prose", async () => {
    let callCount = 0;
    let correctionPromptSeen = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        if (callCount === 1) {
          return buildChatResult(conversationId, payload, "M365_FINAL_V1");
        }
        correctionPromptSeen = readPrompt(payload);
        return buildChatResult(
          conversationId,
          payload,
          "M365_FINAL_V1\nCorrected final text",
        );
      }, SimulatedOutputProtocols.BridgeV1),
    );

    const res = await postResponses(app, {
      model: "m365-copilot",
      stream: false,
      input: "Summarize this tool-free request",
    });
    expect(res.status).toBe(200);
    expect(callCount).toBe(2);
    const body = (await res.json()) as JsonObject;
    expect(body.status).toBe("completed");
    expect(body.output_text).toBe("Corrected final text");
    expect(correctionPromptSeen).toContain("PROTOCOL CORRECTION 1:");
    expect(correctionPromptSeen).toContain(
      "Return either one valid V1 frame starting at byte zero or the strict legacy JSON fallback envelope.",
    );
    expect(correctionPromptSeen).not.toContain(
      "Return one complete JSON object in the fenced format above.",
    );
  });

  test("JSON response format stays legacy under bridge_v1 config", async () => {
    const v1Options = createOptions(SimulatedOutputProtocols.BridgeV1);
    const parsed = tryParseResponsesRequest(
      {
        model: "m365-copilot",
        input: "Return user json",
        response_format: { type: "json_object" },
      },
      v1Options,
    );
    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.simulatedOutputProtocol).toBe(
      SimulatedOutputProtocols.Legacy,
    );
    expect(parsed.request.base.promptText).not.toContain("M365_FINAL_V1");

    const jsonOutput = JSON.stringify({
      object: "response",
      status: "completed",
      output: [
        {
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: '{"success":true}' }],
        },
      ],
    });
    const app = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildChatResult(conversationId, payload, `\`\`\`json\n${jsonOutput}\n\`\`\``),
        SimulatedOutputProtocols.BridgeV1,
      ),
    );
    const res = await postResponses(app, {
      model: "m365-copilot",
      stream: false,
      input: "Return user json",
      response_format: { type: "json_object" },
    });
    expect(res.status).toBe(200);
    const body = (await res.json()) as JsonObject;
    expect(body.status).toBe("completed");
    expect(body.output_text).toBe('{"success":true}');
  });

  test("required and named tool policy rejects V1 final", async () => {
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        if (callCount === 1) {
          return buildChatResult(
            conversationId,
            payload,
            "M365_FINAL_V1\nAttempting to skip tool call",
          );
        }
        return buildChatResult(
          conversationId,
          payload,
          'M365_FUNCTION_TOOL_CALL_V1\nexec\n{"command":"pwd"}',
        );
      }, SimulatedOutputProtocols.BridgeV1),
    );

    const res = await postResponses(app, {
      model: "m365-copilot",
      stream: false,
      input: "Must execute command",
      tools: [functionTool("exec")],
      tool_choice: "required",
    });
    expect(res.status).toBe(200);
    expect(callCount).toBe(2);
    const body = (await res.json()) as JsonObject;
    expect(body.status).toBe("completed");
    const output = body.output as JsonObject[];
    expect(output[0]?.type).toBe("function_call");
    expect(output[0]?.name).toBe("exec");
  });

  test("regeneration and correction prompts remain protocol-aware", () => {
    const originalPrompt = "Request body...\nSTRICT OUTPUT CONTRACT (BRIDGE_V1): ...";

    const v1Regen = buildSimulatedRegenerationPrompt(
      originalPrompt,
      4000,
      SimulatedOutputProtocols.BridgeV1,
    );
    expect(v1Regen).toContain("PROTOCOL REGENERATION:");
    expect(v1Regen).toContain(
      "Return either one valid V1 frame starting at byte zero or the strict legacy JSON fallback envelope.",
    );
    expect(v1Regen).not.toContain("Return exactly one complete JSON object");

    const legacyRegen = buildSimulatedRegenerationPrompt(
      originalPrompt,
      4000,
      SimulatedOutputProtocols.Legacy,
    );
    expect(legacyRegen).toContain("PROTOCOL REGENERATION:");
    expect(legacyRegen).toContain(
      "Return exactly one complete JSON object in the fenced format above.",
    );

    const v1Correction = appendSimulatedProtocolCorrection(
      originalPrompt,
      1,
      "invalid_envelope",
      4000,
      "M365_FINAL_V1",
      SimulatedOutputProtocols.BridgeV1,
    );
    expect(v1Correction).toContain("PROTOCOL CORRECTION 1:");
    expect(v1Correction).toContain(
      "Return either one valid V1 frame starting at byte zero or the strict legacy JSON fallback envelope.",
    );
    expect(v1Correction).not.toContain(
      "Return one complete JSON object in the fenced format above.",
    );

    const legacyCorrection = appendSimulatedProtocolCorrection(
      originalPrompt,
      1,
      "invalid_envelope",
      4000,
      "M365_FINAL_V1",
      SimulatedOutputProtocols.Legacy,
    );
    expect(legacyCorrection).toContain("PROTOCOL CORRECTION 1:");
    expect(legacyCorrection).toContain(
      "Return one complete JSON object in the fenced format above.",
    );
  });
});

function functionTool(name: string): OpenAiToolDefinition {
  return {
    name,
    type: "function",
    description: null,
    parameters: {
      type: "object",
      properties: { command: { type: "string" } },
      required: [],
      additionalProperties: true,
    },
    format: null,
  };
}

function customTool(name: string): OpenAiToolDefinition {
  return {
    name,
    type: "custom",
    description: null,
    parameters: {},
    format: null,
  };
}

function createOptions(protocol?: SimulatedOutputProtocol): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    simulatedOutputProtocol: protocol ?? SimulatedOutputProtocols.Legacy,
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
      maxSendChars: 6000,
      truncateBeforeSending: true,
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}

function createServices(
  onChat: (conversationId: string, payload: JsonObject) => ChatResult,
  protocol?: SimulatedOutputProtocol,
): Parameters<typeof createProxyApp>[0] {
  const options = createOptions(protocol);
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);

  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_test_1",
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
      payload: JsonObject,
    ): Promise<ChatResult> => onChat(conversationId, payload),
    chatOverStream: async (): Promise<Response> => {
      throw new Error("chatOverStream is not used in simulated mode tests.");
    },
  } as unknown as CopilotGraphClient;

  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_substrate_unused",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("substrate chat is not used in this test.");
    },
    chatStream: async (): Promise<ChatResult> => {
      throw new Error("substrate stream is not used in this test.");
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

function buildChatResult(
  conversationId: string,
  payload: JsonObject,
  assistantText: string,
): ChatResult {
  return {
    isSuccess: true,
    statusCode: 200,
    responseJson: {
      id: conversationId,
      messages: [{ text: readPrompt(payload) }, { text: assistantText }],
    },
    rawBody: "{}",
    assistantText: null,
    conversationId,
  };
}

function postResponses(
  app: ReturnType<typeof createProxyApp>,
  body: JsonObject,
  signal?: AbortSignal,
): Promise<Response> {
  return app.fetch(
    new Request("http://localhost/v1/responses", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-m365-transport": TransportNames.Graph,
      },
      body: JSON.stringify(body),
      signal,
    }),
  );
}

function readPrompt(payload: JsonObject): string {
  const message = payload.message;
  if (!message || typeof message !== "object" || Array.isArray(message)) {
    return "";
  }
  const text = (message as Record<string, unknown>).text;
  return typeof text === "string" ? text : "";
}
