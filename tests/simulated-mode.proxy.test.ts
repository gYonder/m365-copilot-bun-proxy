import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  appendSimulatedProtocolCorrection,
  SIMULATED_CORRECTION_RESERVE_CHARS,
  tryParseResponsesRequest,
} from "../src/proxy/request-parser";
import { createProxyApp } from "../src/proxy/server";
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
  type ParsedOpenAiRequest,
  type WrapperOptions,
} from "../src/proxy/types";
import {
  isJsonObject,
  readSseEvents,
  tryGetBoolean,
  tryGetString,
  tryParseJsonObject,
} from "../src/proxy/utils";

describe("simulated transform mode proxy flow", () => {
  test("GET /healthz reports active transform mode", async () => {
    const app = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildGraphChatResult(conversationId, payload, "unused"),
        (options) => {
          options.openAiTransformMode = OpenAiTransformModes.Mapped;
          options.defaultModel = "gpt-5.6-sol";
        },
      ),
    );

    const response = await app.fetch(new Request("http://localhost/healthz"));

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.status).toBe("ok");
    expect(body.openAiTransformMode).toBe(OpenAiTransformModes.Mapped);
    expect(body.defaultModel).toBe("gpt-5.6-sol");
    expect(body.configured_context_limit).toBe(128_000);
    expect(body.observed_safe_context_limit).toBeNull();
    expect(body.verified_provider_context_limit).toBeNull();
  });

  test("mapped responses prompt inlines required local tool contract", () => {
    const options = createOptions();
    options.openAiTransformMode = OpenAiTransformModes.Mapped;

    const parsed = tryParseResponsesRequest(
      {
        model: "m365-copilot",
        stream: false,
        input: "Run pwd using the local shell tool, then answer with the command output.",
        tools: [
          {
            type: "function",
            name: "exec_command",
            description: "Run a shell command.",
            parameters: {
              type: "object",
              properties: {
                cmd: { type: "string" },
                workdir: { type: "string" },
              },
              required: ["cmd"],
            },
          },
        ],
        tool_choice: "auto",
      },
      options,
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) {
      return;
    }
    expect(parsed.request.base.tooling.toolChoiceMode).toBe(
      ToolChoiceModes.Required,
    );
    expect(parsed.request.base.promptText).toContain(
      "A tool call is required for this turn.",
    );
    expect(parsed.request.base.promptText).toContain(
      "Available tool names: exec_command",
    );
    expect(parsed.request.base.promptText).toContain(
      "do not answer from an M365, Python, or /mnt/file_upload environment",
    );
    expect(parsed.request.base.promptText).toContain(
      "If tool output is present, treat that output as the source of truth",
    );
  });

  test("responses reads Codex additional_tools function and custom declarations", () => {
    const parsed = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "additional_tools",
            role: "developer",
            tools: [
              {
                type: "custom",
                name: "apply_patch",
                description: "Apply a patch.",
                format: { type: "grammar", syntax: "lark", definition: "patch" },
              },
              {
                type: "function",
                name: "exec_command",
                parameters: {
                  type: "object",
                  properties: { cmd: { type: "string" } },
                  required: ["cmd"],
                },
              },
            ],
          },
          { type: "message", role: "user", content: "inspect the workspace" },
        ],
      },
      createOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.tooling.tools).toEqual([
      {
        name: "apply_patch",
        type: "custom",
        description: "Apply a patch.",
        parameters: {},
        format: { type: "grammar", syntax: "lark", definition: "patch" },
      },
      {
        name: "exec_command",
        type: "function",
        description: null,
        parameters: {
          type: "object",
          properties: { cmd: { type: "string" } },
          required: ["cmd"],
        },
        format: null,
      },
    ]);
  });

  test("responses reads Codex 0.147 namespace-grouped additional_tools", () => {
    const parsed = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "additional_tools",
            role: "developer",
            tools: [
              {
                type: "namespace",
                name: "functions",
                tools: [
                  {
                    type: "custom",
                    name: "exec",
                    description: "Execute a tool expression.",
                    format: { type: "grammar", syntax: "lark" },
                  },
                ],
              },
            ],
          },
          { type: "message", role: "user", content: "inspect the workspace" },
        ],
      },
      createOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.tooling.tools).toEqual([
      {
        name: "exec",
        type: "custom",
        namespace: "functions",
        description: "Execute a tool expression.",
        parameters: {},
        format: { type: "grammar", syntax: "lark" },
      },
    ]);
  });

  test("responses derives native search only from the latest user message", () => {
    const search = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "message",
            role: "developer",
            content: "The instructions mention web search, current events, and news.",
          },
          {
            type: "message",
            role: "user",
            content: "What is the current top headline on BBC News?",
          },
        ],
      },
      createOptions(),
    );
    expect(search.ok).toBeTrue();
    if (!search.ok) return;
    expect(search.request.base.hostedWebSearch).toBeTrue();

    const coding = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "message",
            role: "developer",
            content: "The instructions mention web search, current events, and news.",
          },
          {
            type: "message",
            role: "user",
            content: "Explain this TypeScript function.",
          },
        ],
      },
      createOptions(),
    );
    expect(coding.ok).toBeTrue();
    if (!coding.ok) return;
    expect(coding.request.base.hostedWebSearch).toBeFalse();
  });

  test(
    "responses compacts oversized tool output without losing turn structure",
    () => {
      const options = createOptions();
      options.substrate.truncateBeforeSending = true;
      options.substrate.maxSendChars = 6_000;
      const toolOutput = `RESULT-HEAD\n${"x".repeat(20_000)}\nRESULT-TAIL`;
      const parsed = tryParseResponsesRequest(
        {
          model: "gpt-5.6-sol",
          previous_response_id: "resp_previous",
          input: [
            {
              type: "function_call_output",
              call_id: "call_oversized",
              output: toolOutput,
            },
            {
              type: "message",
              role: "user",
              content: "Continue using the command evidence.",
            },
          ],
        },
        options,
      );

      expect(parsed.ok).toBeTrue();
      if (!parsed.ok) return;
      const prompt = parsed.request.base.promptText;
      expect(prompt.length).toBeLessThanOrEqual(6_000);
      expect(prompt).toContain('"call_id":"call_oversized"');
      expect(prompt).toContain("RESULT-HEAD");
      expect(prompt).toContain("RESULT-TAIL");
      expect(prompt).toContain(
        "characters omitted from oversized tool result",
      );
      expect(prompt).toContain("Continue using the command evidence.");
      expect(prompt).toContain('"previous_response_id":"resp_previous"');
    },
  );

  test("responses caps one huge tool result below the whole-prompt ceiling", () => {
    const options = createOptions();
    options.substrate.truncateBeforeSending = true;
    options.substrate.maxSendChars = 500_000;
    const parsed = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "function_call_output",
            call_id: "call_huge",
            output: `HEAD-${"x".repeat(520_000)}-TAIL`,
          },
          { type: "message", role: "user", content: "Continue." },
        ],
      },
      options,
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.promptText.length).toBeLessThan(50_000);
    expect(parsed.request.base.promptText).toContain("HEAD-");
    expect(parsed.request.base.promptText).toContain("-TAIL");
    expect(parsed.request.base.promptText).toContain(
      "characters omitted from oversized tool result",
    );
  });

  test("bounds the initial and correction prompts without cutting fenced request JSON", () => {
    const options = createOptions();
    options.substrate.truncateBeforeSending = true;
    options.substrate.maxSendChars = 6_000;
    const parsed = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "function_call_output",
            call_id: "call_bounded",
            output: `HEAD-${"x".repeat(20_000)}-TAIL`,
          },
          { type: "message", role: "user", content: "Continue." },
        ],
      },
      options,
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    const initial = parsed.request.base.promptText;
    const correction = appendSimulatedProtocolCorrection(
      initial,
      1,
      "duplicate_call_id",
      options.substrate.maxSendChars,
    );
    expect(initial.length).toBeLessThanOrEqual(
      options.substrate.maxSendChars - SIMULATED_CORRECTION_RESERVE_CHARS,
    );
    expect(correction.length).toBeLessThanOrEqual(
      options.substrate.maxSendChars,
    );
    const fenced = /```json\n([\s\S]*?)\n```\nSTRICT OUTPUT CONTRACT:/.exec(
      correction,
    );
    expect(fenced).not.toBeNull();
    expect(JSON.parse(fenced?.[1] ?? "")).toBeObject();
    expect(correction).toEndWith(
      "follow the strict output contract above.",
    );
  });

  test("GET /v1/models returns all supported models", async () => {
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(conversationId, payload, "unused"),
      ),
    );

    const response = await app.fetch(new Request("http://localhost/v1/models"));
    const openAiResponse = await app.fetch(
      new Request("http://localhost/openai/v1/models"),
    );

    expect(response.status).toBe(200);
    expect(openAiResponse.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    const openAiBody = (await openAiResponse.json()) as JsonObject;
    expect(tryGetString(body, "object")).toBe("list");
    expect(openAiBody).toEqual(body);
    const data = Array.isArray(body.data) ? (body.data as JsonObject[]) : [];
    const ids = data.map((item) => tryGetString(item, "id"));
    expect(ids).toEqual(["gpt-5.6-sol"]);
    const codexModels = Array.isArray(body.models)
      ? (body.models as JsonObject[])
      : [];
    expect(codexModels.map((item) => tryGetString(item, "slug"))).toEqual(ids);
    expect(codexModels[0]?.shell_type).toBe("shell_command");
    expect(codexModels[0]?.default_reasoning_level).toBe("high");
    expect(codexModels[0]?.supported_reasoning_levels?.map((entry) => entry.effort)).toEqual([
      "none",
      "low",
      "medium",
      "high",
      "xhigh",
      "max",
    ]);
    expect(codexModels[0]?.context_window).toBe(128_000);
    expect(codexModels[0]?.apply_patch_tool_type).toBe("freeform");
    expect(codexModels[0]?.supports_parallel_tool_calls).toBe(true);
    expect(codexModels[0]?.input_modalities).toEqual(["text", "image"]);
  });

  test("chat/completions non-stream wraps incoming JSON and returns parsed JSON block", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_1",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "hello from simulated mode",
          },
          finish_reason: "stop",
        },
      ],
    };

    let capturedPrompt = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompt = readPrompt(payload);
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        );
      }),
    );

    const requestBody: JsonObject = {
      model: "m365-copilot",
      stream: false,
      messages: [{ role: "user", content: "Say hello." }],
    };
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.id).toMatch(/^chatcmpl-/);
    expect((body.choices as JsonObject[])[0]?.message).toEqual(
      (simulatedCompletion.choices as JsonObject[])[0]?.message,
    );
    expect(typeof capturedPrompt).toBe("string");
    expect(capturedPrompt).toContain(
      "The JSON payload below is an entire request for the OpenAI chat.completions format.",
    );
    expect(capturedPrompt).toContain(
      "Focus on producing a valid response object that matches the expected OpenAI format for this request.",
    );
    expect(capturedPrompt).not.toContain("You are simulating");
    expect(capturedPrompt).toContain("```json");
    expect(capturedPrompt).toContain("\"messages\"");
  });

  test("chat/completions non-stream synthesizes local usage and ignores provider usage", async () => {
    const simulatedCompletion: JsonObject = {
      object: "chat.completion",
      choices: [{
        index: 0,
        message: { role: "assistant", content: "usage test" },
        finish_reason: "stop",
      }],
      usage: { prompt_tokens: 999, completion_tokens: 999, total_tokens: 999 },
    };
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        ),
      ),
    );
    const request = {
      model: "m365-copilot",
      stream: false,
      messages: [{ role: "user", content: "Estimate this locally." }],
    };
    const first = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(request),
      }),
    );
    const second = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(request),
      }),
    );

    const firstBody = (await first.json()) as JsonObject;
    const secondBody = (await second.json()) as JsonObject;
    expect(firstBody.usage).toEqual(secondBody.usage);
    expect(firstBody.usage).toMatchObject({
      prompt_tokens: expect.any(Number),
      completion_tokens: expect.any(Number),
      total_tokens: expect.any(Number),
    });
    expect((firstBody.usage as JsonObject).total_tokens).not.toBe(999);
  });

  test("chat/completions stream uses simulated JSON payload", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_stream",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_sim_1",
                type: "function",
                function: {
                  name: "get_time",
                  arguments: "{\"zone\":\"UTC\"}",
                },
              },
            ],
          },
          finish_reason: "tool_calls",
        },
      ],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Call get_time for UTC." }],
          tools: [
            {
              type: "function",
              function: {
                name: "get_time",
                description: "Get the time in a zone.",
                parameters: {
                  type: "object",
                  properties: { zone: { type: "string" } },
                  required: ["zone"],
                  additionalProperties: false,
                },
              },
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();

    let sawToolDelta = false;
    let sawDone = false;
    let finishReason: string | null = null;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const typed = first as Record<string, unknown>;
      if (typeof typed.finish_reason === "string") {
        finishReason = typed.finish_reason;
      }
      const delta = typed.delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const toolCalls = (delta as Record<string, unknown>).tool_calls;
      if (Array.isArray(toolCalls) && toolCalls.length > 0) {
        sawToolDelta = true;
      }
    }

    expect(sawToolDelta).toBeTrue();
    expect(finishReason).toBe("tool_calls");
    expect(sawDone).toBeTrue();
  });

  test("chat/completions stream emits a final usage-only chunk when requested", async () => {
    const simulatedCompletion: JsonObject = {
      object: "chat.completion",
      choices: [{
        index: 0,
        message: { role: "assistant", content: "stream usage" },
        finish_reason: "stop",
      }],
      usage: { prompt_tokens: 999, completion_tokens: 999, total_tokens: 999 },
    };
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        ),
      ),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          stream_options: { include_usage: true },
          messages: [{ role: "user", content: "Stream usage locally." }],
        }),
      }),
    );

    let usageChunk: JsonObject | null = null;
    for await (const event of readSseEvents(response.body!)) {
      if (event.data.trim().toLowerCase() === "[done]") {
        break;
      }
      const chunk = tryParseJsonObject(event.data.trim());
      if (
        chunk &&
        Array.isArray(chunk.choices) &&
        chunk.choices.length === 0 &&
        isJsonObject(chunk.usage)
      ) {
        usageChunk = chunk;
      }
    }

    expect(usageChunk).not.toBeNull();
    expect(usageChunk?.choices).toEqual([]);
    expect(usageChunk?.usage).toMatchObject({
      prompt_tokens: expect.any(Number),
      completion_tokens: expect.any(Number),
      total_tokens: expect.any(Number),
    });
    expect((usageChunk?.usage as JsonObject).total_tokens).not.toBe(999);
  });

  test("chat/completions simulated streaming remains buffered", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_stream_substrate",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "hello from simulated substrate stream",
          },
          finish_reason: "stop",
        },
      ],
    };

    let chatCallCount = 0;
    let chatStreamCallCount = 0;
    const app = createProxyApp(
      createSubstrateStreamingServices(async (onStreamUpdate) => {
        chatStreamCallCount += 1;
        await onStreamUpdate({
          deltaText: toMarkdownJson(simulatedCompletion),
          conversationId: "conv_simulated_substrate_stream",
        });
        return buildGraphChatResult(
          "conv_simulated_substrate_stream",
          {},
          toMarkdownJson(simulatedCompletion),
        );
      }, () => {
        chatCallCount += 1;
        return buildGraphChatResult(
          "conv_simulated_substrate_stream",
          {},
          toMarkdownJson(simulatedCompletion),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();
    expect(chatCallCount).toBe(1);
    expect(chatStreamCallCount).toBe(0);

    let streamedText = "";
    let sawDone = false;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const delta = (first as Record<string, unknown>).delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const content = (delta as Record<string, unknown>).content;
      if (typeof content === "string") {
        streamedText += content;
      }
    }

    expect(streamedText).toContain("hello from simulated substrate stream");
    expect(sawDone).toBeTrue();
  });

  test("chat/completions stream can emit incremental content before full JSON parse when enabled", async () => {
    const finalContent = "hello from incremental simulated streaming";
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_stream_incremental",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: finalContent,
          },
          finish_reason: "stop",
        },
      ],
    };
    const markdownPayload = toMarkdownJson(simulatedCompletion);

    const app = createProxyApp(
      createSubstrateStreamingServices(
        async (onStreamUpdate) => {
          const step = 40;
          for (let i = 0; i < markdownPayload.length; i += step) {
            await onStreamUpdate({
              deltaText: markdownPayload.slice(i, i + step),
              conversationId: "conv_simulated_substrate_incremental",
            });
          }
          return buildGraphChatResult(
            "conv_simulated_substrate_incremental",
            {},
            markdownPayload,
          );
        },
        () =>
          buildGraphChatResult(
            "conv_simulated_substrate_incremental",
            {},
            markdownPayload,
          ),
        (options) => {
          options.substrate.incrementalSimulatedContentStreaming = true;
        },
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();

    const contentChunks: string[] = [];
    let streamedText = "";
    let sawDone = false;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const delta = (first as Record<string, unknown>).delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const content = (delta as Record<string, unknown>).content;
      if (typeof content === "string" && content.length > 0) {
        contentChunks.push(content);
        streamedText += content;
      }
    }

    expect(contentChunks).toHaveLength(1);
    expect(streamedText).toBe(finalContent);
    expect(sawDone).toBeTrue();
  });

  test("chat/completions stream rejects toolless strict-tool payload without resending", async () => {
    const toollessPayload: JsonObject = {
      id: "chatcmpl_simulated_stream_toolless",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "I can help with that.",
          },
          finish_reason: "stop",
        },
      ],
    };
    const toolPayload: JsonObject = {
      id: "chatcmpl_simulated_stream_toolcall",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_readme_1",
                type: "function",
                function: {
                  name: "read_file",
                  arguments: "{\"path\":\"README.md\"}",
                },
              },
            ],
          },
          finish_reason: "tool_calls",
        },
      ],
    };

    let chatCallCount = 0;
    let chatStreamCallCount = 0;
    const app = createProxyApp(
      createSubstrateStreamingServices(async (onStreamUpdate) => {
        chatStreamCallCount += 1;
        await onStreamUpdate({
          deltaText: toMarkdownJson(toollessPayload),
          conversationId: "conv_simulated_substrate_retry",
        });
        return buildGraphChatResult(
          "conv_simulated_substrate_retry",
          {},
          toMarkdownJson(toollessPayload),
        );
      }, () => {
        chatCallCount += 1;
        return buildGraphChatResult(
          "conv_simulated_substrate_retry",
          {},
          toMarkdownJson(toolPayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Read README." }],
          tools: [
            {
              type: "function",
              function: {
                name: "read_file",
                description: "Read a file",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                  },
                  required: ["path"],
                },
              },
            },
          ],
          tool_choice: "required",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();
    expect(chatStreamCallCount).toBe(0);
    expect(chatCallCount).toBe(1);

    let sawToolDelta = false;
    let sawDone = false;
    let finishReason: string | null = null;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const typed = first as Record<string, unknown>;
      if (typeof typed.finish_reason === "string") {
        finishReason = typed.finish_reason;
      }
      const delta = typed.delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const toolCalls = (delta as Record<string, unknown>).tool_calls;
      if (Array.isArray(toolCalls) && toolCalls.length > 0) {
        sawToolDelta = true;
      }
    }

    expect(sawToolDelta).toBeTrue();
    expect(finishReason).toBe("tool_calls");
    expect(sawDone).toBeTrue();
  });

  test("chat/completions stream in simulated mode does not force retry for auto tool choice", async () => {
    const toollessPayload: JsonObject = {
      id: "chatcmpl_simulated_stream_auto_toolless",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "streamed text without forced tool call",
          },
          finish_reason: "stop",
        },
      ],
    };

    let chatCallCount = 0;
    let chatStreamCallCount = 0;
    const app = createProxyApp(
      createSubstrateStreamingServices(async (onStreamUpdate) => {
        chatStreamCallCount += 1;
        await onStreamUpdate({
          deltaText: toMarkdownJson(toollessPayload),
          conversationId: "conv_simulated_substrate_auto_no_retry",
        });
        return buildGraphChatResult(
          "conv_simulated_substrate_auto_no_retry",
          {},
          toMarkdownJson(toollessPayload),
        );
      }, () => {
        chatCallCount += 1;
        return buildGraphChatResult(
          "conv_simulated_substrate_auto_no_retry",
          {},
          toMarkdownJson(toollessPayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Summarize README." }],
          tools: [
            {
              type: "function",
              function: {
                name: "read_file",
                description: "Read a file",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                  },
                  required: ["path"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();
    expect(chatStreamCallCount).toBe(0);
    expect(chatCallCount).toBe(1);

    let streamedText = "";
    let sawDone = false;
    let sawError = false;
    let finishReason: string | null = null;
    for await (const event of readSseEvents(response.body!)) {
      if (event.event === "error") {
        sawError = true;
      }
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const candidateFinishReason = (first as Record<string, unknown>)
        .finish_reason;
      if (typeof candidateFinishReason === "string") {
        finishReason = candidateFinishReason;
      }
      const delta = (first as Record<string, unknown>).delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const content = (delta as Record<string, unknown>).content;
      if (typeof content === "string") {
        streamedText += content;
      }
    }

    expect(streamedText).toContain("streamed text without forced tool call");
    expect(sawDone).toBeTrue();
    expect(sawError).toBeFalse();
    expect(finishReason).toBe("stop");
  });

  test("responses non-stream returns simulated response payload object", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_simulated_1",
      object: "response",
      created_at: 1700000000,
      status: "completed",
      model: "simulated-model",
      output: [
        {
          id: "msg_sim_1",
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: "hello from responses mode" }],
        },
      ],
      output_text: "hello from responses mode",
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const createResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Say hello.",
        }),
      }),
    );

    expect(createResponse.status).toBe(200);
    const body = (await createResponse.json()) as JsonObject;
    expect(body.id).toMatch(/^resp_/);
    expect(body.output_text).toBe("hello from responses mode");
  });

  test("responses preserves validated partial item status in the local projection", async () => {
    const simulatedResponse: JsonObject = {
      object: "response",
      status: "incomplete",
      incomplete_details: { reason: "provider deadline" },
      output: [{
        id: "msg_partial",
        type: "message",
        status: "incomplete",
        role: "assistant",
        content: [{ type: "output_text", text: "partial" }],
      }],
      output_text: "partial",
    };
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          input: "Continue.",
        }),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.status).toBe("incomplete");
    expect((body.output as JsonObject[])[0]?.status).toBe("incomplete");
  });

  test("responses auto tool choice accepts final assistant text without resend", async () => {
    let upstreamCalls = 0;
    const payload: JsonObject = {
      id: "resp_auto_text",
      object: "response",
      status: "completed",
      output: [
        {
          status: "completed",
          type: "message",
          role: "assistant",
          content: [{ type: "output_text", text: "done" }],
        },
      ],
      output_text: "done",
    };
    const app = createProxyApp(
      createServices((conversationId, requestPayload) => {
        upstreamCalls += 1;
        return buildGraphChatResult(
          conversationId,
          requestPayload,
          toMarkdownJson(payload),
        );
      }),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          input: [
            {
              type: "function_call",
              call_id: "call_1",
              name: "noop",
              arguments: "{}",
            },
            {
              type: "function_call_output",
              call_id: "call_1",
              output: "ok",
            },
            { role: "user", content: "Give the final answer." },
          ],
          tools: [
            {
              type: "function",
              name: "noop",
              parameters: { type: "object" },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );
    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.output_text).toBe("done");
    expect(upstreamCalls).toBe(1);
  });

  test("responses stream wraps direct final assistant text without resend", async () => {
    let upstreamCalls = 0;
    const app = createProxyApp(
      createServices((conversationId, requestPayload) => {
        upstreamCalls += 1;
        return buildGraphChatResult(
          conversationId,
          requestPayload,
          toMarkdownJson({
            object: "response",
            status: "completed",
            output: [{
              type: "message",
              status: "completed",
              role: "assistant",
              content: [{ type: "output_text", text: "done" }],
            }],
            output_text: "done",
          }),
        );
      }),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: [
            {
              type: "function_call",
              call_id: "call_1",
              name: "noop",
              arguments: "{}",
            },
            {
              type: "function_call_output",
              call_id: "call_1",
              output: "ok",
            },
            { role: "user", content: "Give the final answer." },
          ],
          tools: [
            {
              type: "function",
              name: "noop",
              parameters: { type: "object" },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );
    expect(response.status).toBe(200);
    let completedOutput = "";
    let sawFailure = false;
    for await (const event of readSseEvents(response.body!)) {
      if (event.event === "response.failed" || event.event === "error") {
        sawFailure = true;
      }
      if (event.event !== "response.completed") continue;
      const parsed = tryParseJsonObject(event.data);
      const completed = parsed?.response;
      if (isJsonObject(completed)) {
        completedOutput = tryGetString(completed, "output_text") ?? "";
      }
    }
    expect(completedOutput).toBe("done");
    expect(sawFailure).toBeFalse();
    expect(upstreamCalls).toBe(1);
  });

  test("responses required tool choice rejects final assistant text without resend", async () => {
    let upstreamCalls = 0;
    const app = createProxyApp(
      createServices((conversationId, requestPayload) => {
        upstreamCalls += 1;
        return buildGraphChatResult(
          conversationId,
          requestPayload,
          toMarkdownJson({
            id: "resp_required_text",
            object: "response",
            status: "completed",
            output: [
              {
                type: "message",
                status: "completed",
                role: "assistant",
                content: [{ type: "output_text", text: "done" }],
              },
            ],
            output_text: "done",
          }),
        );
      }),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          input: "Use the required tool, then answer.",
          tools: [
            {
              type: "function",
              name: "noop",
              parameters: { type: "object" },
            },
          ],
          tool_choice: "required",
        }),
      }),
    );
    expect(response.status).toBe(200);
    expect(upstreamCalls).toBe(2);
    const failure = (await response.json()) as JsonObject;
    expect(failure.status).toBe("failed");
  });

  test("responses non-stream normalizes output_text output items", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_simulated_output_text_item",
      object: "response",
      created_at: 1700000000,
      status: "completed",
      model: "simulated-model",
      output: [{ type: "output_text", text: "hello from output_text item" }],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const createResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Say hello.",
        }),
      }),
    );

    expect(createResponse.status).toBe(200);
    const body = (await createResponse.json()) as JsonObject;
    expect(body.status).toBe("failed");
  });

  test("responses non-stream normalizes assistant role/content output items", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_simulated_role_content",
      object: "response",
      created_at: 1700000000,
      status: "completed",
      model: "simulated-model",
      output: [{ role: "assistant", content: "hello from role/content item" }],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const createResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Say hello.",
        }),
      }),
    );

    expect(createResponse.status).toBe(200);
    const body = (await createResponse.json()) as JsonObject;
    expect(body.status).toBe("failed");
  });

  test("responses non-stream preserves whitespace across output text parts", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_simulated_text_parts",
      object: "response",
      status: "completed",
      output: [{
        type: "message",
        status: "completed",
        role: "assistant",
        content: [
          { type: "output_text", text: "Repository" },
          { type: "output_text", text: " reconnaissance " },
          { type: "output_text", text: "is reconciled." },
        ],
      }],
    };
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Report status.",
        }),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.output_text).toBe(
      "Repository reconnaissance is reconciled.",
    );
  });

  test("responses non-stream rejects outputs aliases at the strict boundary", async () => {
    const simulatedResponse: JsonObject = {
      id: "response-002",
      object: "response",
      model: "m365-copilot",
      outputs: [
        {
          type: "function_call",
          function_call: {
            name: "get_weather",
            arguments: "{\"location\":\"London\",\"unit\":\"celsius\"}",
          },
        },
      ],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const createResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "What's the weather in London?",
          tools: [
            {
              type: "function",
              function: {
                name: "get_weather",
                description: "Get the weather for a location.",
                parameters: {
                  type: "object",
                  properties: {
                    location: { type: "string" },
                    unit: { type: "string" },
                  },
                  required: ["location", "unit"],
                },
              },
            },
          ],
        }),
      }),
    );

    expect(createResponse.status).toBe(200);
    const body = (await createResponse.json()) as JsonObject;
    expect(body.id).toMatch(/^resp_/);
    expect(Array.isArray(body.output)).toBeTrue();
    expect(body.outputs).toBeUndefined();
    expect(body.status).toBe("failed");
    expect(body.output).toEqual([]);
  });

  test("responses stream assigns unique proxy response ids per request", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_001",
      object: "response",
      model: "simulated-model",
      output: [
        {
          id: "msg_simulated_fixed",
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: "hi" }],
        },
      ],
      output_text: "hi",
      status: "completed",
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const collectResponseCreatedId = async (response: Response): Promise<string> => {
      expect(response.status).toBe(200);
      expect(response.body).not.toBeNull();
      for await (const event of readSseEvents(response.body!)) {
        const data = event.data.trim();
        if (!data || data.toLowerCase() === "[done]") {
          continue;
        }
        const parsed = tryParseJsonObject(data);
        if (!parsed) {
          continue;
        }
        if (tryGetString(parsed, "type") !== "response.created") {
          continue;
        }
        if (!isJsonObject(parsed.response)) {
          continue;
        }
        const id = tryGetString(parsed.response as JsonObject, "id");
        if (id) {
          return id;
        }
      }
      throw new Error("missing response.created id");
    };

    const first = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: [{ role: "user", content: [{ type: "input_text", text: "A" }] }],
        }),
      }),
    );
    const second = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: [{ role: "user", content: [{ type: "input_text", text: "B" }] }],
        }),
      }),
    );

    const firstId = await collectResponseCreatedId(first);
    const secondId = await collectResponseCreatedId(second);
    expect(firstId).toContain("resp_");
    expect(secondId).toContain("resp_");
    expect(firstId).not.toBe("resp_001");
    expect(secondId).not.toBe("resp_001");
    expect(firstId).not.toBe(secondId);
  });

  test("responses stream emits canonical in_progress response shape", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_001",
      object: "response",
      created: 1700000000,
      model: "simulated-model",
      output: [
        {
          id: "msg_simulated_shape",
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: "hi" }],
        },
      ],
      output_text: "hi",
      status: "completed",
      reasoning: { encrypted_content: "abc" },
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: [{ role: "user", content: [{ type: "input_text", text: "Hi" }] }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();
    let created: JsonObject | null = null;
    let inProgress: JsonObject | null = null;
    let completed: JsonObject | null = null;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data || data.toLowerCase() === "[done]") {
        if (data.toLowerCase() === "[done]") {
          break;
        }
        continue;
      }
      const parsed = tryParseJsonObject(data);
      if (!parsed) {
        continue;
      }
      const type = tryGetString(parsed, "type");
      if (type === "response.created" && isJsonObject(parsed.response)) {
        created = parsed.response as JsonObject;
      }
      if (type === "response.in_progress" && isJsonObject(parsed.response)) {
        inProgress = parsed.response as JsonObject;
      }
      if (type === "response.completed" && isJsonObject(parsed.response)) {
        completed = parsed.response as JsonObject;
      }
    }

    expect(isJsonObject(created)).toBeTrue();
    expect(isJsonObject(inProgress)).toBeTrue();
    expect(isJsonObject(completed)).toBeTrue();
    const createdResponse = created as JsonObject;
    const inProgressResponse = inProgress as JsonObject;
    const completedResponse = completed as JsonObject;
    expect(tryGetString(createdResponse, "id")?.startsWith("resp_")).toBeTrue();
    expect(tryGetString(inProgressResponse, "id")).toBe(
      tryGetString(createdResponse, "id"),
    );
    expect(tryGetString(completedResponse, "id")).toBe(
      tryGetString(createdResponse, "id"),
    );
    expect(tryGetString(createdResponse, "status")).toBe("in_progress");
    expect(tryGetString(inProgressResponse, "status")).toBe("in_progress");
    expect(tryGetString(completedResponse, "status")).toBe("completed");
    expect(Array.isArray(createdResponse.output)).toBeTrue();
    expect(Array.isArray(inProgressResponse.output)).toBeTrue();
    expect((createdResponse.output as unknown[]).length).toBe(0);
    expect((inProgressResponse.output as unknown[]).length).toBe(0);
    expect(createdResponse.usage).toBeNull();
    expect(inProgressResponse.usage).toBeNull();
    expect(tryGetString(createdResponse, "output_text") ?? "").toBe("");
    expect(tryGetString(inProgressResponse, "output_text") ?? "").toBe("");
    expect(tryGetBoolean(completedResponse, "store")).toBeTrue();
    expect(isJsonObject(completedResponse.usage)).toBeTrue();
    const usage = completedResponse.usage as JsonObject;
    expect(typeof usage.input_tokens).toBe("number");
    expect(typeof usage.output_tokens).toBe("number");
  });

  test.each([
    {
      status: "failed",
      terminalType: "response.failed",
      terminalFields: {
        error: { code: "provider_failed", message: "provider said no" },
        incomplete_details: null,
      },
    },
    {
      status: "incomplete",
      terminalType: "response.incomplete",
      terminalFields: {
        error: null,
        incomplete_details: { reason: "provider deadline" },
      },
    },
  ])("replays a $status Responses stream with a clean in_progress snapshot", async ({
    status,
    terminalType,
    terminalFields,
  }) => {
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "provider-response-id",
            object: "response",
            model: "provider-model",
            created_at: 1,
            status,
            output: [],
            output_text: "",
            usage: { total_tokens: 999 },
            ...terminalFields,
          }),
        ),
      ),
    );
    const requestBody = {
      model: "m365-copilot",
      stream: true,
      client_metadata: {
        thread_id: `thread-${status}`,
        session_id: `session-${status}`,
        turn_id: `turn-${status}`,
      },
      input: "Return the provider terminal state.",
    };
    const send = () =>
      app.fetch(
        new Request("http://localhost/v1/responses", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-m365-transport": TransportNames.Graph,
          },
          body: JSON.stringify(requestBody),
        }),
      );
    const collect = async (response: Response): Promise<JsonObject[]> => {
      const events: JsonObject[] = [];
      for await (const event of readSseEvents(response.body!)) {
        const data = event.data.trim();
        if (!data || data === "[DONE]") continue;
        const parsed = tryParseJsonObject(data);
        if (parsed) events.push(parsed);
      }
      return events;
    };

    const firstEvents = await collect(await send());
    const replayEvents = await collect(await send());
    const inProgress = replayEvents.find(
      (event) => event.type === "response.in_progress",
    )?.response as JsonObject;
    const terminal = replayEvents.find(
      (event) => event.type === terminalType,
    );

    expect(inProgress).toMatchObject({
      status: "in_progress",
      error: null,
      incomplete_details: null,
      usage: null,
      output: [],
    });
    expect(terminal).toMatchObject({
      type: terminalType,
      response: {
        status,
        ...terminalFields,
        usage: expect.any(Object),
        output: [],
      },
    });
    const terminalResponse = terminal?.response as JsonObject;
    expect(tryGetString(terminalResponse, "id")?.startsWith("resp_")).toBeTrue();
    expect(tryGetString(terminalResponse, "model")).toBe("m365-copilot");
    expect(terminalResponse.created_at).not.toBe(1);
    expect(
      (terminalResponse.usage as JsonObject).total_tokens,
    ).not.toBe(999);
    expect(firstEvents.some((event) => event.type === terminalType)).toBeTrue();
    expect(replayEvents.at(-1)?.type).toBe(terminalType);
  });

  test("responses accepts spec conversation string input and returns spec conversation output", async () => {
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_spec_conversation_string",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "message",
                role: "assistant",
                content: [{ type: "output_text", text: "hello from conversation string" }],
              },
            ],
            output_text: "hello from conversation string",
          }),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          conversation: "conv_spec_string_1",
          input: "Say hello.",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.headers.get("x-m365-conversation-id")).toBe("conv_spec_string_1");
    const body = (await response.json()) as JsonObject;
    expect(tryGetString(body, "conversation")).toBe("conv_spec_string_1");
    expect(tryGetString(body, "conversation_id")).toBe("conv_spec_string_1");
  });

  test("responses accepts spec conversation object input", async () => {
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_spec_conversation_object",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "message",
                role: "assistant",
                content: [{ type: "output_text", text: "hello from conversation object" }],
              },
            ],
            output_text: "hello from conversation object",
          }),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          conversation: { id: "conv_spec_object_1" },
          input: "Say hello.",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.headers.get("x-m365-conversation-id")).toBe("conv_spec_object_1");
    const body = (await response.json()) as JsonObject;
    expect(tryGetString(body, "conversation")).toBe("conv_spec_object_1");
    expect(tryGetString(body, "conversation_id")).toBe("conv_spec_object_1");
  });

  test("responses rejects conversation combined with previous_response_id", async () => {
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_unused",
            object: "response",
            output: [],
            output_text: "",
          }),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          conversation: "conv_spec_string_2",
          previous_response_id: "resp_prev_1",
          input: "Say hello.",
        }),
      }),
    );

    expect(response.status).toBe(400);
    const body = (await response.json()) as JsonObject;
    expect(isJsonObject(body.error)).toBeTrue();
    expect(tryGetString(body.error as JsonObject, "code")).toBe("invalid_request");
  });

  test("responses continuation advances once and replays an identical retry", async () => {
    let upstreamCalls = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        upstreamCalls += 1;
        const response = {
          id: `resp_tool_${upstreamCalls}`,
          object: "response",
          status: "completed",
          output: [
            {
              status: "completed",
              type: "function_call",
              call_id: "call_1",
              name: "exec",
              arguments: "{}",
            },
          ],
        };
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(response),
        );
      }),
    );
    const headers = {
      "content-type": "application/json",
      "x-m365-transport": TransportNames.Graph,
    };
    const first = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers,
        body: JSON.stringify({
          model: "m365-copilot",
          input: "Use exec.",
          tools: [
            {
              type: "function",
              name: "exec",
              parameters: { type: "object" },
            },
          ],
          tool_choice: "required",
        }),
      }),
    );
    expect(first.status).toBe(200);
    const firstBody = (await first.json()) as JsonObject;
    const firstId = tryGetString(firstBody, "id");
    expect(firstId).not.toBeNull();
    const continuationBody = {
      model: "m365-copilot",
      previous_response_id: firstId,
      input: [
        {
          type: "function_call",
          call_id: "call_1",
          name: "exec",
          arguments: "{}",
        },
        { type: "function_call_output", call_id: "call_1", output: "ok" },
        {
          type: "custom_tool_call_output",
          call_id: "custom_1",
          output: "ok",
        },
        { role: "user", content: "Give final answer." },
      ],
      tools: [
        {
          type: "function",
          name: "exec",
          parameters: { type: "object" },
        },
      ],
      tool_choice: "auto",
    };
    const second = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers,
        body: JSON.stringify(continuationBody),
      }),
    );
    expect(second.status).toBe(200);
    const secondBody = (await second.json()) as JsonObject;
    const secondOutput = Array.isArray(secondBody.output)
      ? secondBody.output
      : [];
    expect(
      secondOutput.some(
        (item) =>
          isJsonObject(item) &&
          tryGetString(item, "type") === "function_call",
      ),
    ).toBeTrue();

    const third = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers,
        body: JSON.stringify(continuationBody),
      }),
    );
    expect(third.status).toBe(200);
    expect(third.headers.get("x-m365-request-hash-replayed")).toBeNull();
    expect(third.headers.get("x-m365-protocol-identity-replayed")).toBe("true");
    const thirdBody = (await third.json()) as JsonObject;
    expect(thirdBody).toEqual(secondBody);
    expect(upstreamCalls).toBe(2);
  });

  test("concurrent duplicate protocol turns use one upstream invocation and preserve one tool call", async () => {
    let upstreamCalls = 0;
    let releaseUpstream: (() => void) | undefined;
    const upstreamBlocked = new Promise<void>((resolve) => {
      releaseUpstream = resolve;
    });
    const app = createProxyApp(
      createServices(async (conversationId, payload) => {
        upstreamCalls += 1;
        await upstreamBlocked;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_concurrent_tool",
            object: "response",
            status: "completed",
            output: [
              {
                status: "completed",
                type: "function_call",
                call_id: "call_concurrent_1",
                name: "exec",
                arguments: JSON.stringify({ cmd: "pwd" }),
              },
            ],
          }),
        );
      }),
    );
    const body = JSON.stringify({
      model: "m365-copilot",
      stream: false,
      input: "Run pwd once.",
      client_metadata: {
        thread_id: "thread-concurrent",
        session_id: "session-concurrent",
        turn_id: "turn-concurrent",
      },
      tools: [
        {
          type: "function",
          name: "exec",
          parameters: {
            type: "object",
            properties: { cmd: { type: "string" } },
            required: ["cmd"],
          },
        },
      ],
      tool_choice: "required",
    });
    const makeRequest = () =>
      app.fetch(
        new Request("http://localhost/v1/responses", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-m365-transport": TransportNames.Graph,
          },
          body,
        }),
      );

    const firstPromise = makeRequest();
    const secondPromise = makeRequest();
    await Promise.resolve();
    releaseUpstream?.();
    const [first, second] = await Promise.all([firstPromise, secondPromise]);

    expect(first.status).toBe(200);
    expect(second.status).toBe(200);
    expect(upstreamCalls).toBe(1);
    expect(
      [first, second].filter(
        (response) =>
          response.headers.get("x-m365-in-flight-replayed") === "true",
      ),
    ).toHaveLength(1);
    const firstBody = (await first.json()) as JsonObject;
    const secondBody = (await second.json()) as JsonObject;
    expect(secondBody).toEqual(firstBody);
    const output = Array.isArray(firstBody.output) ? firstBody.output : [];
    expect(
      output.filter(
        (item) =>
          isJsonObject(item) &&
          tryGetString(item, "type") === "function_call" &&
          tryGetString(item, "call_id") === "call_concurrent_1",
      ),
    ).toHaveLength(1);
  });

  test("responses request-hash guard replays the stored response for duplicate identical requests", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_hash_guard_first",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                status: "completed",
                type: "message",
                role: "assistant",
                content: [{ type: "output_text", text: "hello once" }],
              },
            ],
            output_text: "hello once",
          }),
        );
      }),
    );

    const requestBody = {
      model: "m365-copilot",
      stream: false,
      input: [{ role: "user", content: [{ type: "input_text", text: "Hi" }] }],
    };

    const first = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );
    expect(first.status).toBe(200);
    const firstBody = (await first.json()) as JsonObject;
    expect(tryGetString(firstBody, "output_text")).toBe("hello once");

    const second = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );

    expect(second.status).toBe(200);
    // Upstream is still hit only once — the guard prevents the duplicate from
    // re-invoking the substrate conversation.
    expect(chatCallCount).toBe(1);
    expect(second.headers.get("x-m365-request-hash-replayed")).toBe("true");
    expect(second.headers.get("x-m365-replay-suppressed")).toBe("true");
    // New: the duplicate now replays the REAL stored body (idempotent) rather
    // than an empty message, so retried tool turns keep their output.
    expect(second.headers.get("x-m365-replay-idempotent")).toBe("true");
    expect(second.headers.get("x-m365-conversation-id")).toBe("conv_simulated_1");
    const secondBody = (await second.json()) as JsonObject;
    expect(tryGetString(secondBody, "status")).toBe("completed");
    expect(tryGetString(secondBody, "id")).toBe(
      tryGetString(firstBody, "id"),
    );
    expect(Array.isArray(secondBody.output)).toBeTrue();
    expect((secondBody.output as unknown[]).length).toBe(1);
    expect(tryGetString(secondBody, "output_text") ?? "").toBe("hello once");
  });

  test("responses request-hash guard replays custom patch calls on duplicate", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_tool_first",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "custom_tool_call",
                status: "completed",
                call_id: "call_abc123",
                name: "apply_patch",
                input: "*** Begin Patch\n*** Add File: hello.txt\n+world\n*** End Patch",
              },
            ],
          }),
        );
      }),
    );

    const requestBody = {
      model: "m365-copilot",
      stream: false,
      input: [
        {
          role: "user",
          content: [{ type: "input_text", text: "Create hello.txt with world" }],
        },
      ],
      tools: [
        {
          type: "custom",
          name: "apply_patch",
          format: { type: "grammar", syntax: "lark", definition: "patch" },
        },
      ],
      tool_choice: "auto",
    };

    const makeRequest = () =>
      app.fetch(
        new Request("http://localhost/v1/responses", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-m365-transport": TransportNames.Graph,
          },
          body: JSON.stringify(requestBody),
        }),
      );

    const first = await makeRequest();
    expect(first.status).toBe(200);
    const firstBody = (await first.json()) as JsonObject;
    const firstOutput = firstBody.output as JsonObject[];
    expect(firstOutput[0]?.type).toBe("custom_tool_call");

    const second = await makeRequest();
    expect(second.status).toBe(200);
    expect(chatCallCount).toBe(1);
    expect(second.headers.get("x-m365-replay-idempotent")).toBe("true");

    const secondBody = (await second.json()) as JsonObject;
    const secondOutput = secondBody.output as JsonObject[];
    // The retried tool turn MUST still carry the freeform patch so Codex can
    // apply the patch / run the shell command — not an empty message.
    expect(Array.isArray(secondOutput)).toBeTrue();
    expect(secondOutput[0]?.type).toBe("custom_tool_call");
    expect(tryGetString(secondOutput[0], "name")).toBe("apply_patch");
    expect(tryGetString(secondOutput[0], "call_id")).toBe("call_abc123");
    expect(tryGetString(secondOutput[0], "input")).toBe(
      "*** Begin Patch\n*** Add File: hello.txt\n+world\n*** End Patch",
    );
  });

  test("does not replay identical bodies across explicit conversation headers", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: `resp-${conversationId}`,
            object: "response",
            status: "completed",
            model: "simulated-model",
            output: [
              {
                status: "completed",
                type: "message",
                role: "assistant",
                content: [{ type: "output_text", text: conversationId }],
              },
            ],
            output_text: conversationId,
          }),
        );
      }),
    );
    const body = JSON.stringify({
      model: "m365-copilot",
      stream: false,
      input: "same body",
    });
    const send = (conversationId: string) =>
      app.fetch(
        new Request("http://localhost/v1/responses", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-m365-transport": TransportNames.Graph,
            "x-m365-conversation-id": conversationId,
          },
          body,
        }),
      );

    const first = await send("conversation-a");
    const second = await send("conversation-b");

    expect(first.status).toBe(200);
    expect(second.status).toBe(200);
    expect(chatCallCount).toBe(2);
    expect(second.headers.get("x-m365-request-hash-replayed")).toBeNull();
    expect(tryGetString((await second.json()) as JsonObject, "output_text")).toBe(
      "conversation-b",
    );
  });

  test("streams the full custom tool-call input lifecycle in order with a stable item id", async () => {
    const expectedPatch =
      "\n*** Begin Patch\n*** Update File: probe.txt\n@@\n-hello\n+world\n*** End Patch\n";
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_custom_stream",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "custom_tool_call",
                status: "completed",
                call_id: "call_stream_patch",
                name: "apply_patch",
                input: expectedPatch,
              },
            ],
          }),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          conversation: "conv_custom_stream",
          input: [
            {
              role: "user",
              content: [
                {
                  type: "input_text",
                  text: "Change probe.txt from hello to world using apply_patch.",
                },
              ],
            },
          ],
          tools: [
            {
              type: "custom",
              name: "apply_patch",
              format: { type: "grammar", syntax: "lark", definition: "patch" },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(
      (response.headers.get("content-type") ?? "").includes("text/event-stream"),
    ).toBeTrue();
    expect(response.body).not.toBeNull();

    const orderedTypes: string[] = [];
    let addedItemId: string | null = null;
    let deltaItemId: string | null = null;
    let doneItemId: string | null = null;
    let itemDoneId: string | null = null;
    let deltaText: string | null = null;
    let doneInput: string | null = null;
    let completedResponse: JsonObject | null = null;

    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        break;
      }
      const parsed = tryParseJsonObject(data);
      if (!parsed) {
        continue;
      }
      const type = tryGetString(parsed, "type");
      if (!type) {
        continue;
      }
      orderedTypes.push(type);
      if (type === "response.output_item.added") {
        const item = isJsonObject(parsed.item) ? (parsed.item as JsonObject) : null;
        if (item && tryGetString(item, "type") === "custom_tool_call") {
          addedItemId = tryGetString(item, "id");
          expect(tryGetString(item, "call_id")).toBe("call_stream_patch");
          expect(tryGetString(item, "name")).toBe("apply_patch");
        }
      }
      if (type === "response.custom_tool_call_input.delta") {
        deltaItemId = tryGetString(parsed, "item_id");
        deltaText = typeof parsed.delta === "string" ? parsed.delta : null;
      }
      if (type === "response.custom_tool_call_input.done") {
        doneItemId = tryGetString(parsed, "item_id");
        doneInput = typeof parsed.input === "string" ? parsed.input : null;
      }
      if (type === "response.output_item.done") {
        const item = isJsonObject(parsed.item) ? (parsed.item as JsonObject) : null;
        if (item && tryGetString(item, "type") === "custom_tool_call") {
          itemDoneId = tryGetString(item, "id");
        }
      }
      if (type === "response.completed" && isJsonObject(parsed.response)) {
        completedResponse = parsed.response as JsonObject;
      }
    }

    // Exact custom tool-call streaming lifecycle order (handover item 4). The
    // filter keeps only the lifecycle-defining events so the comparison is
    // insensitive to unrelated keep-alive or text events but strict on order.
    const lifecycle = orderedTypes.filter((eventType) =>
      [
        "response.created",
        "response.in_progress",
        "response.output_item.added",
        "response.custom_tool_call_input.delta",
        "response.custom_tool_call_input.done",
        "response.output_item.done",
        "response.completed",
      ].includes(eventType),
    );
    expect(lifecycle).toEqual([
      "response.created",
      "response.in_progress",
      "response.output_item.added",
      "response.custom_tool_call_input.delta",
      "response.custom_tool_call_input.done",
      "response.output_item.done",
      "response.completed",
    ]);

    // The item id must stay stable across added -> delta -> done -> item.done so
    // Codex can correlate the custom tool-call input stream to one output item.
    expect(addedItemId).toBeTruthy();
    expect(deltaItemId).toBe(addedItemId);
    expect(doneItemId).toBe(addedItemId);
    expect(itemDoneId).toBe(addedItemId);

    // Freeform patch bytes are preserved verbatim through delta and done.
    expect(deltaText).toBe(expectedPatch);
    expect(doneInput).toBe(expectedPatch);

    // The completed response carries the correlated custom tool call.
    expect(isJsonObject(completedResponse)).toBeTrue();
    const output = (completedResponse as JsonObject).output as JsonObject[];
    expect(output[0]?.type).toBe("custom_tool_call");
    expect(tryGetString(output[0], "call_id")).toBe("call_stream_patch");
    expect(tryGetString(output[0], "name")).toBe("apply_patch");
    expect(output[0]?.input).toBe(expectedPatch);
  });

  test("reconstructs synthesized Responses SSE payloads without duplicated parts or arguments", async () => {
    const expectedPatch = "*** Begin Patch\n*** Add File: exact.txt\n+exact\n*** End Patch";
    const providerResponse: JsonObject = {
      object: "response",
      status: "completed",
      output: [
        {
          id: "provider_message",
          type: "message",
          status: "completed",
          role: "assistant",
          content: [
            { type: "output_text", text: "first " },
            { type: "refusal", refusal: "not this part" },
            { type: "output_text", text: "second" },
          ],
        },
        {
          id: "provider_function",
          type: "function_call",
          status: "completed",
          call_id: "call_function",
          name: "exec",
          arguments: '{"cmd":"pwd"}',
        },
        {
          id: "provider_custom",
          type: "custom_tool_call",
          status: "completed",
          call_id: "call_custom",
          name: "apply_patch",
          input: expectedPatch,
        },
      ],
      output_text: "first second",
    };
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(providerResponse),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: "Return the exact mixed response.",
          tools: [
            {
              type: "function",
              name: "exec",
              parameters: { type: "object" },
            },
            {
              type: "custom",
              name: "apply_patch",
              format: {
                type: "grammar",
                syntax: "lark",
                definition: "start: /[\\s\\S]+/",
              },
            },
          ],
          tool_choice: "auto",
          parallel_tool_calls: true,
        }),
      }),
    );

    const events: JsonObject[] = [];
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data || data === "[DONE]") continue;
      const parsed = tryParseJsonObject(data);
      if (parsed) events.push(parsed);
    }

    const added = events.filter((event) =>
      event.type === "response.output_item.added"
    );
    expect(added).toHaveLength(3);
    for (const event of added) {
      const item = event.item as JsonObject;
      expect(item.status).toBe("in_progress");
      if (item.type === "message") {
        expect(item.content).toEqual([]);
      } else if (item.type === "function_call") {
        expect(item.arguments).toBe("");
      } else if (item.type === "custom_tool_call") {
        expect(item.input).toBe("");
      }
    }

    const reconstructed = new Map<number, JsonObject>();
    const completedItems = new Map<number, JsonObject>();
    for (const event of added) {
      const index = Number(event.output_index);
      reconstructed.set(index, {
        ...(event.item as JsonObject),
        content: event.item && (event.item as JsonObject).type === "message"
          ? []
          : (event.item as JsonObject).content,
      });
    }
    for (const event of events) {
      const index = Number(event.output_index);
      const item = reconstructed.get(index);
      if (!item) continue;
      if (event.type === "response.content_part.added") {
        const content = Array.isArray(item.content) ? item.content : [];
        content[Number(event.content_index)] = event.part;
        item.content = content;
      } else if (event.type === "response.content_part.done") {
        const content = Array.isArray(item.content) ? item.content : [];
        content[Number(event.content_index)] = event.part;
        item.content = content;
      } else if (event.type === "response.output_text.delta") {
        const content = item.content as JsonObject[];
        const part = content[Number(event.content_index)] as JsonObject;
        part.text = `${part.text ?? ""}${event.delta ?? ""}`;
      } else if (event.type === "response.refusal.delta") {
        const content = item.content as JsonObject[];
        const part = content[Number(event.content_index)] as JsonObject;
        part.refusal = `${part.refusal ?? ""}${event.delta ?? ""}`;
      } else if (event.type === "response.function_call_arguments.delta") {
        item.arguments = `${item.arguments ?? ""}${event.delta ?? ""}`;
      } else if (event.type === "response.custom_tool_call_input.delta") {
        item.input = `${item.input ?? ""}${event.delta ?? ""}`;
      } else if (event.type === "response.output_item.done") {
        completedItems.set(index, event.item as JsonObject);
      }
    }

    for (const [index, item] of completedItems) {
      const reconstructedItem = reconstructed.get(index)!;
      reconstructedItem.status = item.status;
      expect(reconstructedItem).toEqual(item);
    }
    const completed = events.find((event) => event.type === "response.completed");
    expect(completed).toBeTruthy();
    const completedOutput = (completed!.response as JsonObject).output as JsonObject[];
    expect([...completedItems.values()]).toEqual(completedOutput);
    expect(
      events
        .filter((event) => event.type === "response.output_text.delta")
        .map((event) => event.content_index),
    ).toEqual([0, 2]);
    expect(
      events
        .filter((event) => event.type === "response.refusal.delta")
        .map((event) => event.content_index),
    ).toEqual([1]);
  });


  test("does not infer replay from a trailing assistant tail", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_should_not_be_used",
            object: "response",
            status: "completed",
            output: [{
              type: "message",
              status: "completed",
              role: "assistant",
              content: [{ type: "output_text", text: "unused" }],
            }],
            output_text: "unused",
          }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          conversation: "conv_replay_guard_non_stream",
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(chatCallCount).toBe(1);
    expect(response.headers.get("x-m365-replay-suppressed")).toBeNull();
    expect(response.headers.get("x-m365-conversation-id")).toBe(
      "conv_replay_guard_non_stream",
    );
    const body = (await response.json()) as JsonObject;
    expect(tryGetString(body, "object")).toBe("response");
    expect(tryGetString(body, "status")).toBe("completed");
    expect(tryGetString(body, "conversation")).toBe(
      "conv_replay_guard_non_stream",
    );
    expect(tryGetString(body, "conversation_id")).toBe(
      "conv_replay_guard_non_stream",
    );
    expect(Array.isArray(body.output)).toBeTrue();
    expect((body.output as unknown[]).length).toBe(1);
    expect(tryGetString(body, "output_text") ?? "").toBe("unused");
  });

  test("does not infer streaming replay from a trailing assistant tail", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_should_not_be_used",
            object: "response",
            status: "completed",
            output: [{
              type: "message",
              status: "completed",
              role: "assistant",
              content: [{ type: "output_text", text: "unused" }],
            }],
            output_text: "unused",
          }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          conversation: "conv_replay_guard_stream",
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(chatCallCount).toBe(1);
    expect(response.headers.get("x-m365-replay-suppressed")).toBeNull();
    expect(response.headers.get("x-m365-conversation-id")).toBe(
      "conv_replay_guard_stream",
    );
    const contentType = response.headers.get("content-type") ?? "";
    expect(contentType.includes("text/event-stream")).toBeTrue();
    expect(response.body).not.toBeNull();

    const eventTypes: string[] = [];
    const sseEventNames: string[] = [];
    let completedResponse: JsonObject | null = null;
    let sawDone = false;
    for await (const event of readSseEvents(response.body!)) {
      sseEventNames.push(event.event);
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const parsed = tryParseJsonObject(data);
      if (!parsed) {
        continue;
      }
      const type = tryGetString(parsed, "type");
      if (!type) {
        continue;
      }
      eventTypes.push(type);
      if (type === "response.completed" && isJsonObject(parsed.response)) {
        completedResponse = parsed.response as JsonObject;
      }
    }

    expect(eventTypes).toContain("response.created");
    expect(eventTypes).toContain("response.in_progress");
    expect(eventTypes).toContain("response.completed");
    expect(sseEventNames).toContain("response.created");
    expect(sseEventNames).toContain("response.in_progress");
    expect(sseEventNames).toContain("response.completed");
    expect(sseEventNames).toContain("done");
    expect(isJsonObject(completedResponse)).toBeTrue();
    const completed = completedResponse as JsonObject;
    expect(tryGetString(completed, "status")).toBe("completed");
    expect(tryGetString(completed, "conversation")).toBe(
      "conv_replay_guard_stream",
    );
    expect(tryGetString(completed, "conversation_id")).toBe(
      "conv_replay_guard_stream",
    );
    expect(Array.isArray(completed.output)).toBeTrue();
    expect((completed.output as unknown[]).length).toBe(1);
    expect(tryGetString(completed, "output_text") ?? "").toBe("unused");
    expect(sawDone).toBeTrue();
  });

  test("does not infer replay from one assistant item", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_should_not_be_used_single_assistant",
            object: "response",
            status: "completed",
            output: [{
              type: "message",
              status: "completed",
              role: "assistant",
              content: [{ type: "output_text", text: "unused" }],
            }],
            output_text: "unused",
          }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(chatCallCount).toBe(1);
    expect(response.headers.get("x-m365-replay-suppressed")).toBeNull();
    const body = (await response.json()) as JsonObject;
    expect(tryGetString(body, "status")).toBe("completed");
    expect(Array.isArray(body.output)).toBeTrue();
    expect((body.output as unknown[]).length).toBe(1);
    expect(tryGetString(body, "output_text") ?? "").toBe("unused");
  });

  test("keeps explicit conversation header without content-tail replay", async () => {
    const app = createProxyApp(
      createServices(
        (conversationId, payload) =>
          buildGraphChatResult(
            conversationId,
            payload,
            toMarkdownJson({
              id: "resp_should_not_be_used_header_only",
              object: "response",
              status: "completed",
              output: [{
                type: "message",
                status: "completed",
                role: "assistant",
                content: [{ type: "output_text", text: "unused" }],
              }],
              output_text: "unused",
            }),
          ),
        (options) => {
          options.includeConversationIdInResponseBody = false;
        },
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          conversation: "conv_header_only_replay",
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.headers.get("x-m365-replay-suppressed")).toBeNull();
    expect(response.headers.get("x-m365-conversation-id")).toBe(
      "conv_header_only_replay",
    );
    const body = (await response.json()) as JsonObject;
    expect(tryGetString(body, "conversation")).toBeNull();
    expect(tryGetString(body, "conversation_id")).toBeNull();
    expect(tryGetString(body, "output_text") ?? "").toBe("unused");
  });

  test("continues growing assistant tails unless request identity repeats", async () => {
    let chatCallCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        chatCallCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_canonical_hash_seed",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                status: "completed",
                type: "message",
                role: "assistant",
                content: [{ type: "output_text", text: "Hi" }],
              },
            ],
            output_text: "Hi",
          }),
        );
      }),
    );

    const first = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );
    expect(first.status).toBe(200);
    expect(chatCallCount).toBe(1);

    const replayOne = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );
    expect(replayOne.status).toBe(200);
    expect(chatCallCount).toBe(2);
    expect(replayOne.headers.get("x-m365-replay-suppressed")).toBeNull();
    expect(replayOne.headers.get("x-m365-conversation-id")).toBe("conv_simulated_1");
    const replayOneBody = (await replayOne.json()) as JsonObject;
    const replayOneId = tryGetString(replayOneBody, "id");
    expect(replayOneId?.startsWith("resp_")).toBeTrue();
    expect(tryGetString(replayOneBody, "output_text") ?? "").toBe("Hi");

    const replayTwo = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              role: "user",
              content: [{ type: "input_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
            {
              role: "assistant",
              content: [{ type: "output_text", text: "Hi" }],
            },
          ],
        }),
      }),
    );
    expect(replayTwo.status).toBe(200);
    expect(chatCallCount).toBe(3);
    expect(replayTwo.headers.get("x-m365-replay-suppressed")).toBeNull();
    expect(replayTwo.headers.get("x-m365-conversation-id")).toBe("conv_simulated_1");
    const replayTwoBody = (await replayTwo.json()) as JsonObject;
    expect(tryGetString(replayTwoBody, "id")?.startsWith("resp_")).toBeTrue();
    expect(tryGetString(replayTwoBody, "id")).not.toBe(replayOneId);
    expect(tryGetString(replayTwoBody, "output_text") ?? "").toBe("Hi");
  });

  test("chat/completions rejects top-level choice-shaped payloads", async () => {
    const malformedChoiceShape: JsonObject = {
      index: 0,
      finish_reason: "tool_calls",
      message: {
        role: "assistant",
        tool_calls: [
          {
            id: "attempt-final-001",
            type: "function",
            function: {
              name: "attempt_completion",
              arguments: {
                result: "done",
              },
            },
          },
        ],
      },
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(malformedChoiceShape),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Complete the task." }],
          tools: [
            {
              type: "function",
              function: {
                name: "attempt_completion",
                description: "Complete the task.",
                parameters: {
                  type: "object",
                  properties: { result: { type: "string" } },
                  required: ["result"],
                  additionalProperties: false,
                },
              },
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(502);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions rejects non-string tool-call arguments", async () => {
    const payloadWithObjectArguments: JsonObject = {
      id: "chatcmpl_obj_args",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            tool_calls: [
              {
                id: "call_write",
                type: "function",
                function: {
                  name: "write_to_file",
                  arguments: {
                    path: "tests/agent-tests/fibonacci.ts",
                    content: "export const x = 1;",
                  },
                },
              },
            ],
          },
        },
      ],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(payloadWithObjectArguments),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Write the file." }],
          tools: [
            {
              type: "function",
              function: {
                name: "write_to_file",
                description: "Write file content.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    content: { type: "string" },
                  },
                  required: ["path", "content"],
                  additionalProperties: false,
                },
              },
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(502);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions simulated prompt includes explicit tool-call guidance", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_prompt_guidance",
      object: "chat.completion",
      model: "simulated-model",
      created: 1700000000,
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_1",
                type: "function",
                function: {
                  name: "write_to_file",
                  arguments: "{\"path\":\"tests/agent-tests/fizz-buzz.ts\",\"content\":\"x\"}",
                },
              },
            ],
          },
        },
      ],
    };

    let capturedPrompt = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompt = readPrompt(payload);
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Implement fizz buzz." }],
          tools: [
            {
              type: "function",
              function: {
                name: "write_to_file",
                description: "Write content",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    content: { type: "string" },
                  },
                  required: ["path", "content"],
                },
              },
            },
          ],
          tool_choice: "required",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(capturedPrompt).toContain(
      "Use choices[0].message.tool_calls",
    );
    expect(capturedPrompt).toContain(
      "If the request requires local files, shell state, or any other local environment access, emit an appropriate tool call instead of saying the environment is inaccessible.",
    );
    expect(capturedPrompt).toContain(
      "Do not claim you inspected, changed, or verified local files unless the response includes the matching tool call.",
    );
    expect(capturedPrompt).toContain(
      "Do not refuse by saying tool invocation is unsupported.",
    );
    expect(capturedPrompt).toContain(
      "function.arguments must be a JSON string value",
    );
    expect(capturedPrompt).toContain(
      "This request requires at least one tool call.",
    );
  });

  test("responses simulated prompt requires tool calls for initial local file tasks", async () => {
    let capturedPrompt = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompt = readPrompt(payload);
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_tool_required",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "function_call",
                call_id: "call_1",
                name: "exec_command",
                arguments: "{\"cmd\":\"cat seed.md\"}",
              },
            ],
          }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              type: "message",
              role: "user",
              content: [
                {
                  type: "input_text",
                  text: "Use local shell/file tools. Read seed.md and write result.md.",
                },
              ],
            },
          ],
          tools: [
            {
              type: "function",
              name: "exec_command",
              description: "Run a command",
              parameters: {
                type: "object",
                properties: { cmd: { type: "string" } },
                required: ["cmd"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(capturedPrompt).toContain(
      "This request requires at least one tool call. Do not return a plain-text-only assistant response.",
    );
    expect(capturedPrompt).toContain(
      "STRICT OUTPUT CONTRACT:",
    );
    expect(capturedPrompt).toContain(
      '"type":"function_call","status":"completed"',
    );
    expect(capturedPrompt.lastIndexOf("STRICT OUTPUT CONTRACT")).toBeGreaterThan(
      capturedPrompt.lastIndexOf("```"),
    );
  });

  test("responses retries an invalid initial local-tool payload with a protocol correction", async () => {
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const services = createSubstrateStreamingServices(
      async () => {
        throw new Error("buffered simulated responses should not use chatStream");
      },
      (request) => {
        callCount += 1;
        capturedPrompts.push(request?.promptText ?? "");
        const payload =
          callCount === 1
            ? {
                id: "resp_toolless_first",
                object: "response",
                status: "completed",
                output: [
                  {
                    type: "message",
                    status: "completed",
                    role: "assistant",
                    content: [{ type: "output_text", text: "The path is unknown." }],
                  },
                ],
                output_text: "The path is unknown.",
              }
            : {
                id: "resp_tool_retry",
                object: "response",
                status: "completed",
                output: [
                  {
                    type: "function_call",
                    status: "completed",
                    call_id: "call_retry",
                    name: "exec_command",
                    arguments: "{\"cmd\":\"pwd\"}",
                  },
                ],
              };
        return buildGraphChatResult(
          "conv_simulated_substrate_stream",
          {},
          toMarkdownJson(payload),
        );
      },
    );
    services.options.substrate.earlyCompleteOnSimulatedPayload = true;
    const app = createProxyApp(services);

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: "Use the shell to run pwd.",
          tools: [
            {
              type: "function",
              name: "exec_command",
              description: "Run a command",
              parameters: {
                type: "object",
                properties: { cmd: { type: "string" } },
                required: ["cmd"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[1]).toContain("PROTOCOL CORRECTION 1:");
    expect(capturedPrompts[1]).toContain(
      "Use {\"object\":\"response\",\"status\":\"completed\"",
    );
  });

  test("responses repairs a malformed full tool envelope instead of exposing it as text", async () => {
    const malformedEnvelope = `{
      "object": "response",
      "status": "completed",
      "output": [{
        "type": "function_call",
        "call_id": "call_malformed_exec",
        "name": "exec",
        "arguments": "const result = await tools.exec_command({
          cmd: "pwd"
        });
        text(result);"
      }]
    }`;
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        capturedPrompts.push(readPrompt(payload));
        return buildGraphChatResult(
          conversationId,
          payload,
          callCount === 1
            ? malformedEnvelope
            : toMarkdownJson({
                id: "resp_repaired_exec",
                object: "response",
                status: "completed",
                output: [
                  {
                    type: "custom_tool_call",
                    status: "completed",
                    call_id: "call_repaired_exec",
                    name: "exec",
                    input:
                      'const result = await tools.exec_command({cmd:"pwd"}); text(result);',
                  },
                ],
              }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Use the local shell to run pwd.",
          tools: [
            {
              type: "custom",
              name: "exec",
              description: "Execute a tool expression.",
              format: {
                type: "grammar",
                syntax: "lark",
                definition: "start: /[\\s\\S]+/",
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[1]).toContain("PROTOCOL CORRECTION 1:");
    const body = (await response.json()) as JsonObject;
    const output = Array.isArray(body.output) ? body.output : [];
    expect(output).toHaveLength(1);
    expect(output[0]).toMatchObject({
      type: "custom_tool_call",
      call_id: "call_repaired_exec",
      name: "exec",
    });
    expect(tryGetString(output[0], "input")).toContain("tools.exec_command");
  });

  test.each([
    {
      label: "fenced",
      malformed: "```json\n{\n  \"id\": \"resp_retry1\",\n  \"object\": \"response\",\n  \"status\": \"in_progress\",\n  \"model\": \"gpt-5.6-sol\",\n  \"output\": const applied = await tools.apply_patch(patch); text(applied);\n}\n```",
    },
    {
      label: "unfenced",
      malformed: "{\"id\":\"resp_retry2\",\"object\":\"response\",\"status\":\"in_progress\",\"model\":\"gpt-5.6-sol\",\"output\":const result = await tools.exec_command({cmd:\"pwd\"});text(JSON.stringify(result));}",
    },
  ])("responses repairs a $label malformed in-progress Responses wrapper", async ({ malformed }) => {
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        capturedPrompts.push(readPrompt(payload));
        return buildGraphChatResult(
          conversationId,
          payload,
          callCount === 1
            ? malformed
            : toMarkdownJson({
                id: "resp_repaired_wrapper",
                object: "response",
                status: "completed",
                output: [{
                  type: "custom_tool_call",
                status: "completed",
                call_id: "call_repaired_wrapper",
                  name: "exec",
                  input: 'const result = await tools.exec_command({cmd:"pwd"}); text(result);',
                }],
              }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Use the local shell to run pwd.",
          tools: [{
            type: "custom",
            name: "exec",
            description: "Execute a tool expression.",
            format: { type: "grammar", syntax: "lark", definition: "start: /[\\s\\S]+/" },
          }],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[1]).toContain("PROTOCOL CORRECTION 1:");
    const body = (await response.json()) as JsonObject;
    const output = Array.isArray(body.output) ? body.output : [];
    expect(output[0]).toMatchObject({
      type: "custom_tool_call",
      call_id: "call_repaired_wrapper",
      name: "exec",
    });
  });

  test("responses rejects a repeated malformed in-progress wrapper after one correction", async () => {
    const malformed = "{\"id\":\"resp_retry_exhausted\",\"object\":\"response\",\"status\":\"in_progress\",\"output\":const result = await tools.exec_command({cmd:\"pwd\"});text(result);}";
    let callCount = 0;
    const app = createProxyApp(createServices((conversationId, payload) => {
      callCount += 1;
      return buildGraphChatResult(conversationId, payload, malformed);
    }));
    const response = await app.fetch(new Request("http://localhost/v1/responses", {
      method: "POST",
      headers: { "content-type": "application/json", "x-m365-transport": TransportNames.Graph },
      body: JSON.stringify({
        model: "m365-copilot",
        stream: false,
        input: "Use the local shell to run pwd.",
        tools: [{ type: "custom", name: "exec", description: "Execute a tool expression.", format: { type: "grammar", syntax: "lark" } }],
        tool_choice: "auto",
      }),
    }));
    expect(callCount).toBe(2);
    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.status).toBe("failed");
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("responses resumes a preceding leaked tool envelope on explicit request", async () => {
    const malformedEnvelope = `{
      "object": "response",
      "output": [{
        "type": "function_call",
        "name": "exec",
        "arguments": "const result = await tools.exec_command({
          cmd: "pwd"
        }); text(result);"
      }]
    }`;
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        const prompt = readPrompt(payload);
        capturedPrompts.push(prompt);
        const assistantText = callCount === 1
          ? malformedEnvelope
          : prompt.includes("RECOVERY TURN")
            ? toMarkdownJson({
                id: "resp_resumed_exec",
                object: "response",
                status: "completed",
                output: [
                  {
                    type: "custom_tool_call",
                    status: "completed",
                    call_id: "call_resumed_exec",
                    name: "exec",
                    input:
                      'const result = await tools.exec_command({cmd:"pwd"}); text(result);',
                  },
                ],
              })
            : "I will continue from where I stopped.";
        return buildGraphChatResult(
          conversationId,
          payload,
          assistantText,
        );
      }),
    );

    const leaked = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Begin the implementation.",
        }),
      }),
    );
    expect(leaked.status).toBe(200);
    const leakedBody = (await leaked.json()) as JsonObject;
    expect(leakedBody.status).toBe("failed");

    const resumed = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          previous_response_id: leakedBody.id,
          input: [
            {
              type: "additional_tools",
              role: "developer",
              tools: [
                {
                  type: "namespace",
                  name: "functions",
                  tools: [
                    {
                      type: "custom",
                      name: "exec",
                      description: "Execute a tool expression.",
                      format: {
                        type: "grammar",
                        syntax: "lark",
                        definition: "start: /[\\s\\S]+/",
                      },
                    },
                  ],
                },
              ],
            },
            {
              type: "message",
              status: "completed",
              role: "user",
              content: [{ type: "input_text", text: "resume work" }],
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(resumed.status).toBe(200);
    expect(callCount).toBe(4);
    expect(capturedPrompts[1]).not.toContain("RECOVERY TURN");
    expect(resumed.headers.get("x-m365-tool-call-recovery")).toBeNull();
    const resumedBody = (await resumed.json()) as JsonObject;
    expect(resumedBody.status).toBe("failed");

    const fullHistoryResume = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: [
            {
              type: "message",
              role: "assistant",
              content: [{ type: "output_text", text: malformedEnvelope }],
            },
            {
              type: "message",
              role: "user",
              content: [{ type: "input_text", text: "continue task" }],
            },
          ],
          tools: [
            {
              type: "custom",
              name: "exec",
              description: "Execute a tool expression.",
              format: { type: "grammar", syntax: "lark" },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(fullHistoryResume.status).toBe(200);
    expect(callCount).toBe(5);
    expect(capturedPrompts[4]).toContain("RECOVERY TURN");
    expect(fullHistoryResume.headers.get("x-m365-tool-call-recovery")).toBe(
      "true",
    );
  });

  test("responses does not force recovery after ordinary assistant text", async () => {
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        capturedPrompts.push(readPrompt(payload));
        return buildGraphChatResult(
          conversationId,
          payload,
          callCount === 1
            ? toMarkdownJson({
                object: "response",
                status: "completed",
                output: [{
                  type: "message",
                  status: "completed",
                  role: "assistant",
                  content: [{ type: "output_text", text: "The requested analysis is complete." }],
                }],
                output_text: "The requested analysis is complete.",
              })
            : toMarkdownJson({
                object: "response",
                status: "completed",
                output: [{
                  type: "message",
                  status: "completed",
                  role: "assistant",
                  content: [{ type: "output_text", text: "Continuing with the next explanation." }],
                }],
                output_text: "Continuing with the next explanation.",
              }),
        );
      }),
    );

    const first = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Explain the current design.",
        }),
      }),
    );
    const firstBody = (await first.json()) as JsonObject;

    const resumed = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          previous_response_id: firstBody.id,
          input: "resume work",
          tools: [
            {
              type: "function",
              name: "exec_command",
              description: "Run a command.",
              parameters: {
                type: "object",
                properties: { cmd: { type: "string" } },
                required: ["cmd"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(resumed.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[1]).not.toContain("RECOVERY TURN");
    expect(resumed.headers.get("x-m365-tool-call-recovery")).toBeNull();
  });

  test("responses bounds buffered simulated recovery to one correction turn", async () => {
    let callCount = 0;
    const services = createSubstrateStreamingServices(
      async () => {
        throw new Error("buffered simulated responses should not use chatStream");
      },
      (_request) => {
        callCount += 1;
        return buildGraphChatResult(
          "conv_latency_bound_" + callCount,
          {},
          toMarkdownJson({
            id: "resp_latency_bound_" + callCount,
            object: "response",
            status: "completed",
            output: [
              {
                type: "message",
                status: "completed",
                role: "assistant",
                content: [{ type: "output_text", text: "I cannot access the tool." }],
              },
            ],
            output_text: "I cannot access the tool.",
          }),
        );
      },
    );
    services.options.confabRetries = 1;
    services.options.substrate.earlyCompleteOnSimulatedPayload = true;
    const app = createProxyApp(services);

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: "Use the shell to run pwd.",
          tools: [
            {
              type: "function",
              name: "exec_command",
              description: "Run a command",
              parameters: {
                type: "object",
                properties: { cmd: { type: "string" } },
                required: ["cmd"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    const events: JsonObject[] = [];
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data || data === "[DONE]") continue;
      const parsed = tryParseJsonObject(data);
      if (parsed) events.push(parsed);
    }
    const terminal = events.find((event) => event.type === "response.failed");
    expect(terminal).toBeTruthy();
    expect((terminal?.response as JsonObject).status).toBe("failed");
  });

  test("responses simulated prompt treats copying previous output to the clipboard as a local tool action", async () => {
    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompts.push(readPrompt(payload));
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(
            callCount === 1
              ? {
                  id: "resp_clipboard_text_only",
                  object: "response",
                  status: "completed",
                  output: [
                    {
                      type: "message",
                      role: "assistant",
                      content: [{ type: "output_text", text: "Copied it." }],
                    },
                  ],
                  output_text: "Copied it.",
                }
              : {
                  id: "resp_clipboard_tool",
                  object: "response",
                  status: "completed",
                  output: [
                    {
                      type: "function_call",
                      status: "completed",
                      call_id: "call_clipboard",
                      name: "exec_command",
                      arguments: "{\"cmd\":\"pbcopy < output.txt\"}",
                    },
                  ],
                },
          ),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Place the entire previous output in my clipboard.",
          tools: [
            {
              type: "function",
              name: "exec_command",
              description: "Run a command",
              parameters: {
                type: "object",
                properties: { cmd: { type: "string" } },
                required: ["cmd"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[0]).toContain(
      "This request requires at least one tool call.",
    );
    const body = (await response.json()) as JsonObject;
    const output = Array.isArray(body.output) ? body.output : [];
    expect(output[0]).toMatchObject({
      type: "function_call",
      call_id: "call_clipboard",
      name: "exec_command",
    });
  });

  test("responses simulated prompt preserves mcp-like tool names", async () => {
    let capturedPrompt = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompt = readPrompt(payload);
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson({
            id: "resp_mcp_like_tool",
            object: "response",
            created_at: 1700000000,
            status: "completed",
            model: "simulated-model",
            output: [
              {
                type: "function_call",
                status: "completed",
                call_id: "call_mcp_1",
                name: "mcp__wiki__search",
                arguments: "{\"query\":\"order allocation\"}",
              },
            ],
          }),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Search the local wiki for order allocation notes.",
          tools: [
            {
              type: "function",
              name: "mcp__wiki__search",
              description: "Search the local companion wiki",
              parameters: {
                type: "object",
                properties: { query: { type: "string" } },
                required: ["query"],
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    const body = await response.json();
    const output = Array.isArray(body.output) ? body.output : [];
    const functionCall = output.find(
      (item) =>
        isJsonObject(item) &&
        tryGetString(item, "type") === "function_call" &&
        tryGetString(item, "name") === "mcp__wiki__search",
    );

    expect(response.status).toBe(200);
    expect(capturedPrompt).toContain('\"name\": \"mcp__wiki__search\"');
    expect(capturedPrompt).toContain(
      "Function calls use",
    );
    expect(functionCall).toBeTruthy();
  });

  test("chat/completions repairs malformed tool-call arguments with raw newlines", async () => {
    const brokenArguments =
      "{\"path\":\"tests/agent-tests/fizz-buzz.ts\",\"diff\":\"<<<<<<< SEARCH\n:start_line:1\nfoo\n=======\nbar\n>>>>>>> REPLACE\"}";
    const payloadWithBrokenArguments: JsonObject = {
      id: "chatcmpl_broken_args",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            tool_calls: [
              {
                id: "call_apply_diff",
                type: "function",
                function: {
                  name: "apply_diff",
                  arguments: brokenArguments,
                },
              },
            ],
          },
        },
      ],
    };

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) =>
        (() => {
          callCount += 1;
          return buildGraphChatResult(
            conversationId,
            payload,
            toMarkdownJson(payloadWithBrokenArguments),
          );
        })(),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Add comments." }],
          tools: [
            {
              type: "function",
              function: {
                name: "apply_diff",
                description: "Apply an exact textual diff.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    diff: { type: "string" },
                  },
                  required: ["path", "diff"],
                  additionalProperties: false,
                },
              },
            },
          ],
        }),
      }),
    );

    expect(response.status).toBe(502);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions rejects an empty simulated payload without resending", async () => {
    const emptyPayload: JsonObject = {
      id: "chatcmpl-empty",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "stop",
          message: { role: "assistant", content: "" },
        },
      ],
    };
    const usablePayload: JsonObject = {
      id: "chatcmpl-usable",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_retry",
                type: "function",
                function: {
                  name: "write_to_file",
                  arguments:
                    "{\"path\":\"tests/agent-tests/fizz-buzz.ts\",\"content\":\"export const ok = true;\"}",
                },
              },
            ],
          },
        },
      ],
    };

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(callCount === 1 ? emptyPayload : usablePayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Implement fizz buzz." }],
        }),
      }),
    );

    expect(response.status).toBe(502);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions accepts a toolless auto-tool payload without resending", async () => {
    const plainTextStopPayload: JsonObject = {
      id: "chatcmpl-plain-stop",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "stop",
          message: {
            role: "assistant",
            content: "I wrote the file.",
          },
        },
      ],
    };
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(plainTextStopPayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Implement fizz buzz." }],
          tools: [
            {
              type: "function",
              function: {
                name: "write_to_file",
                description: "Write file content.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    content: { type: "string" },
                  },
                  required: ["path", "content"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(1);
    const body = (await response.json()) as JsonObject;
    const choices = body.choices as JsonObject[];
    expect(choices[0]?.finish_reason).toBe("stop");
    const message = choices[0]?.message as JsonObject;
    expect(message.content).toBe("I wrote the file.");
  });

  test("chat/completions corrects invalid apply_diff once before delivery", async () => {
    const invalidApplyDiffPayload: JsonObject = {
      id: "chatcmpl-invalid-applydiff",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_apply_diff_bad",
                type: "function",
                function: {
                  name: "apply_diff",
                  arguments:
                    "{\"path\":\"tests/agent-tests/fibonacci.ts\",\"diff\":\"<<<<<<< SEARCH\\n:start_line:1\\n-------\\n\\n=======\\nexport function fibonacci(n: number): number {\\n  if (n <= 1) return n;\\n  return n;\\n}\\n>>>>>>> REPLACE\"}",
                },
              },
            ],
          },
        },
      ],
    };
    const usablePayload: JsonObject = {
      id: "chatcmpl-usable-write",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_write_file_good",
                type: "function",
                function: {
                  name: "write_to_file",
                  arguments:
                    "{\"path\":\"tests/agent-tests/fibonacci.ts\",\"content\":\"export function fibonacci(n: number): number {\\n  if (n <= 1) return n;\\n  return n;\\n}\"}",
                },
              },
            ],
          },
        },
      ],
    };

    const capturedPrompts: string[] = [];
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        capturedPrompts.push(readPrompt(payload));
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(
            callCount === 1 ? invalidApplyDiffPayload : usablePayload,
          ),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Implement fibonacci." }],
          tools: [
            {
              type: "function",
              function: {
                name: "apply_diff",
                description: "Apply textual diff.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    diff: { type: "string" },
                  },
                  required: ["path", "diff"],
                },
              },
            },
            {
              type: "function",
              function: {
                name: "write_to_file",
                description: "Write file content.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    content: { type: "string" },
                  },
                  required: ["path", "content"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    expect(capturedPrompts[1]).toContain("PROTOCOL CORRECTION 1:");
    const body = (await response.json()) as JsonObject;
    const choices = body.choices as JsonObject[];
    const message = choices[0]?.message as JsonObject;
    const calls = message.tool_calls as JsonObject[];
    expect(calls).toHaveLength(1);
    expect(((calls[0]?.function as JsonObject).name)).toBe("write_to_file");
  });

  test("chat/completions corrects duplicate call IDs before projection", async () => {
    const duplicate: JsonObject = {
      object: "chat.completion",
      choices: [{
        index: 0,
        finish_reason: "tool_calls",
        message: {
          role: "assistant",
          content: null,
          tool_calls: [
            {
              id: "call_duplicate",
              type: "function",
              function: { name: "exec", arguments: "{}" },
            },
            {
              id: "call_duplicate",
              type: "function",
              function: { name: "exec", arguments: "{}" },
            },
          ],
        },
      }],
    };
    const corrected: JsonObject = {
      object: "chat.completion",
      choices: [{
        index: 0,
        finish_reason: "tool_calls",
        message: {
          role: "assistant",
          content: null,
          tool_calls: [{
            id: "call_corrected",
            type: "function",
            function: { name: "exec", arguments: "{}" },
          }],
        },
      }],
    };
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(callCount === 1 ? duplicate : corrected),
        );
      }),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          messages: [{ role: "user", content: "Run exec." }],
          tools: [{
            type: "function",
            function: {
              name: "exec",
              parameters: { type: "object" },
            },
          }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    const choices = body.choices as JsonObject[];
    const calls = (choices[0]?.message as JsonObject).tool_calls as JsonObject[];
    expect(calls).toHaveLength(1);
    expect(calls[0]?.id).toBe("call_corrected");
  });

  test("responses corrects duplicate call IDs before ledger mutation", async () => {
    const duplicate: JsonObject = {
      object: "response",
      status: "completed",
      output: [
        responseFunctionCall("exec", "{}", "call_duplicate"),
        responseFunctionCall("exec", "{}", "call_duplicate"),
      ],
    };
    const corrected: JsonObject = {
      object: "response",
      status: "completed",
      output: [responseFunctionCall("exec", "{}", "call_corrected")],
    };
    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(callCount === 1 ? duplicate : corrected),
        );
      }),
    );
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          input: [{ type: "message", role: "user", content: "Run exec." }],
          tools: [{
            type: "function",
            name: "exec",
            parameters: { type: "object" },
          }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    const output = body.output as JsonObject[];
    expect(output).toHaveLength(1);
    expect(output[0]?.call_id).toBe("call_corrected");
  });

  test("does not replay a cancelled Chat correction", async () => {
    const controller = new AbortController();
    let abortNext = true;
    let callCount = 0;
    const invalid = "{\"not\":\"a chat completion\"}";
    const valid = toMarkdownJson({
      object: "chat.completion",
      choices: [{
        index: 0,
        finish_reason: "stop",
        message: { role: "assistant", content: "retried" },
      }],
    });
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        if (abortNext) {
          abortNext = false;
          controller.abort();
          return buildGraphChatResult(conversationId, payload, invalid);
        }
        return buildGraphChatResult(conversationId, payload, valid);
      }),
    );
    const requestBody = {
      model: "m365-copilot",
      messages: [{ role: "user", content: "Retry safely." }],
    };
    const first = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        signal: controller.signal,
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );
    expect(first.status).toBe(499);
    expect((await first.json()).error.code).toBe("client_aborted");

    const second = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );
    expect(second.status).toBe(200);
    expect(callCount).toBe(2);
  });

  test("does not register a cancelled Responses correction for replay", async () => {
    const controller = new AbortController();
    let callCount = 0;
    const invalid = "{\"not\":\"a response\"}";
    const valid = toMarkdownJson({
      object: "response",
      status: "completed",
      output: [responseMessage("retried")],
    });
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        if (callCount === 2) {
          controller.abort();
        }
        return buildGraphChatResult(
          conversationId,
          payload,
          callCount === 1 ? invalid : valid,
        );
      }),
    );
    const requestBody = {
      model: "m365-copilot",
      input: [{ type: "message", role: "user", content: "Retry safely." }],
    };
    const first = await app.fetch(
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
    const firstBody = (await first.json()) as JsonObject;
    expect(firstBody.status).toBe("failed");
    expect((firstBody.error as JsonObject).code).toBe("upstream_timeout");

    const second = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );
    expect(second.status).toBe(200);
    expect((await second.json()).status).toBe("completed");
    expect(callCount).toBe(3);
  });

  test("chat/completions stream accepts toolless auto-tool payload without resending", async () => {
    const plainTextStopPayload: JsonObject = {
      id: "chatcmpl-plain-stop-stream",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "stop",
          message: {
            role: "assistant",
            content: "High-level summary: this proxy maps OpenAI-style traffic to M365 Copilot backends.",
          },
        },
      ],
    };

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(plainTextStopPayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Explain the project." }],
          tools: [
            {
              type: "function",
              function: {
                name: "ask_followup_question",
                description: "Ask a follow-up question.",
                parameters: {
                  type: "object",
                  properties: {
                    question: { type: "string" },
                  },
                  required: ["question"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();
    expect(callCount).toBe(1);

    let streamedText = "";
    let finishReason: string | null = null;
    let sawDone = false;
    let sawError = false;
    for await (const event of readSseEvents(response.body!)) {
      if (event.event === "error") {
        sawError = true;
      }
      const data = event.data.trim();
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || !isJsonObject(choices[0])) {
        continue;
      }
      const choice = choices[0];
      const candidateFinishReason = tryGetString(choice, "finish_reason");
      if (candidateFinishReason) {
        finishReason = candidateFinishReason;
      }
      const delta = choice.delta;
      if (isJsonObject(delta)) {
        streamedText += tryGetString(delta, "content") ?? "";
      }
    }

    expect(streamedText).toContain("High-level summary");
    expect(finishReason).toBe("stop");
    expect(sawDone).toBeTrue();
    expect(sawError).toBeFalse();
  });

  test("chat/completions exhausts one invalid apply_diff correction", async () => {
    const invalidApplyDiffPayload: JsonObject = {
      id: "chatcmpl-invalid-applydiff-persistent",
      object: "chat.completion",
      model: "simulated-model",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_apply_diff_bad_persistent",
                type: "function",
                function: {
                  name: "apply_diff",
                  arguments:
                    "{\"path\":\"tests/agent-tests/fibonacci.ts\",\"diff\":\"<<<<<<< SEARCH\\n:start_line:1\\n-------\\n\\n=======\\nexport function fibonacci(n: number): number {\\n  if (n <= 1) return n;\\n  return n;\\n}\\n>>>>>>> REPLACE\"}",
                },
              },
            ],
          },
        },
      ],
    };

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(invalidApplyDiffPayload),
        );
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Implement fibonacci." }],
          tools: [
            {
              type: "function",
              function: {
                name: "apply_diff",
                description: "Apply textual diff.",
                parameters: {
                  type: "object",
                  properties: {
                    path: { type: "string" },
                    diff: { type: "string" },
                  },
                  required: ["path", "diff"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    expect(response.status).toBe(502);
    expect(callCount).toBe(2);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions returns 502 for a bare tool-call envelope emitted as text after retry", async () => {
    // The model emits a Responses-API-shaped tool call as prose instead of a
    // structured tool call. classifyToolAttempt should detect it as an
    // invalid_attempt, trigger one protocol retry, and if the retry also fails,
    // return 502 instead of passing the envelope through as text content.
    const bareEnvelope = `{"type":"function_call","name":"functions.exec","call_id":"call_bare_1","arguments":"const r = await tools.exec_command({cmd:\\"ls\\"}); text(r);"}`;

    let callCount = 0;
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        callCount += 1;
        return buildGraphChatResult(conversationId, payload, bareEnvelope);
      }),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [
            { role: "user", content: "Run pwd using the local shell tool." },
          ],
          tools: [
            {
              type: "function",
              function: {
                name: "exec_command",
                description: "Run a shell command.",
                parameters: {
                  type: "object",
                  properties: {
                    cmd: { type: "string" },
                  },
                  required: ["cmd"],
                },
              },
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    // Should have retried once (initial + 1 protocol retry)
    expect(callCount).toBe(2);
    // Should return 502, not 200 with the envelope as content
    expect(response.status).toBe(502);
    const body = (await response.json()) as JsonObject;
    expect((body.error as JsonObject).code).toBe("provider_drift");
  });

  test("chat/completions extracts a bare functions.exec envelope when exec is offered as a custom tool", async () => {
    // The model emits functions.exec as text, but the offered tool is named
    // "exec" (custom type). The functions. prefix should be stripped and the
    // call extracted as a valid custom tool call. Uses Mapped mode so the
    // buildAssistantResponse path (where tryBuildToolCall runs) is reached
    // directly, without the simulated-payload extraction gate.
    const bareEnvelope = `{"type":"function_call","name":"functions.exec","call_id":"call_exec_1","arguments":"const r = await tools.exec_command({cmd:\\"pwd\\"}); text(r);"}`;

    let callCount = 0;
    const app = createProxyApp(
      createServices(
        (conversationId, payload) => {
          callCount += 1;
          return buildGraphChatResult(conversationId, payload, bareEnvelope);
        },
        (options) => {
          options.openAiTransformMode = OpenAiTransformModes.Mapped;
        },
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          messages: [{ role: "user", content: "Run pwd." }],
          tools: [
            {
              type: "custom",
              name: "exec",
              description: "Execute a tool expression.",
            },
          ],
          tool_choice: "auto",
        }),
      }),
    );

    // Should succeed on first attempt (no retry needed in Mapped mode)
    expect(callCount).toBe(1);
    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    const choices = body.choices as JsonObject[];
    expect(choices).toHaveLength(1);
    const toolCalls = (choices[0].message as JsonObject).tool_calls as JsonObject[];
    expect(toolCalls).toHaveLength(1);
    expect((toolCalls[0].function as JsonObject).name).toBe("exec");
    // The JS expression should be preserved in arguments
    const args = (toolCalls[0].function as JsonObject).arguments as string;
    expect(args).toContain("exec_command");
    expect(args).toContain("pwd");
    expect(choices[0].finish_reason).toBe("tool_calls");
  });
});

function responseMessage(content: string): JsonObject {
  return {
    id: "msg_test",
    type: "message",
    status: "completed",
    role: "assistant",
    content: [{ type: "output_text", text: content }],
  };
}

function responseFunctionCall(
  name: string,
  argumentsJson: string,
  callId: string,
): JsonObject {
  return {
    id: `item_${callId}`,
    type: "function_call",
    status: "completed",
    call_id: callId,
    name,
    arguments: argumentsJson,
  };
}

function createServices(
  onChat: (conversationId: string, payload: JsonObject) => ChatResult,
  configureOptions?: (options: WrapperOptions) => void,
): Parameters<typeof createProxyApp>[0] {
  const options = createOptions();
  if (configureOptions) {
    configureOptions(options);
  }
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);

  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_simulated_1",
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
    resolveAuthorizationHeader: async () => "Bearer unit-test-token",
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

function createSubstrateStreamingServices(
  onChatStream: (
    onStreamUpdate: (update: {
      deltaText: string | null;
      conversationId: string | null;
    }) => Promise<void>,
  ) => Promise<ChatResult>,
  onChat: (request?: ParsedOpenAiRequest) => ChatResult,
  configureOptions?: (options: WrapperOptions) => void,
): Parameters<typeof createProxyApp>[0] {
  const options = createOptions();
  options.transport = TransportNames.Substrate;
  if (configureOptions) {
    configureOptions(options);
  }
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);

  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_graph_unused",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("graph chat is not used in this test.");
    },
    chatOverStream: async (): Promise<Response> => {
      throw new Error("graph stream is not used in this test.");
    },
  } as unknown as CopilotGraphClient;

  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_simulated_substrate_stream",
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      _conversationId: string,
      request: ParsedOpenAiRequest,
    ): Promise<ChatResult> => onChat(request),
    chatStream: async (
      _authorizationHeader: string,
      _conversationId: string,
      _request: unknown,
      _isStartOfSession: boolean,
      onStreamUpdate: (update: {
        deltaText: string | null;
        conversationId: string | null;
      }) => Promise<void>,
    ): Promise<ChatResult> => onChatStream(onStreamUpdate),
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
    logLevel: LogLevels.Info,
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

function buildGraphChatResult(
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
    conversationId: conversationId,
  };
}

function readPrompt(payload: JsonObject): string {
  const message = payload.message;
  if (!message || typeof message !== "object" || Array.isArray(message)) {
    return "";
  }
  const text = (message as Record<string, unknown>).text;
  return typeof text === "string" ? text : "";
}

function toMarkdownJson(payload: JsonObject): string {
  return `\`\`\`json\n${JSON.stringify(payload, null, 2)}\n\`\`\``;
}
