import { describe, expect, test } from "bun:test";
import { buildFunctionCallOutputItem } from "../src/proxy/responses-api";
import {
  tryParseOpenAiRequest,
  tryParseResponsesRequest,
} from "../src/proxy/request-parser";
import {
  LogLevels,
  OpenAiTransformModes,
  ToolChoiceModes,
  TransportNames,
  type WrapperOptions,
} from "../src/proxy/types";

describe("Responses custom tool support", () => {
  test("emits a custom_tool_call with byte-preserved freeform input", () => {
    const item = buildFunctionCallOutputItem(
      "item_patch",
      { id: "call_patch", name: "apply_patch", type: "custom", argumentsJson: "*** Begin Patch\n*** End Patch" },
      "completed",
    );
    expect(item).toEqual({
      id: "item_patch",
      type: "custom_tool_call",
      status: "completed",
      call_id: "call_patch",
      name: "apply_patch",
      input: "*** Begin Patch\n*** End Patch",
    });
  });

  test("normalizes custom tool call results as correlated tool messages", () => {
    const parsed = tryParseResponsesRequest({
      model: "gpt-5.6-sol",
      input: [{ type: "custom_tool_call_output", call_id: "call_patch", output: "done" }],
    }, options());
    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.inputItemsForStorage).toEqual([
      { type: "custom_tool_call_output", call_id: "call_patch", output: "done" },
    ]);
  });

  test("does not force another initial tool call after custom output", () => {
    const parsed = tryParseResponsesRequest(
      {
        model: "gpt-5.6-sol",
        input: [
          {
            type: "custom_tool_call_output",
            call_id: "call_patch",
            output: "done",
          },
          { role: "user", content: "Give the final answer." },
        ],
        tools: [
          {
            type: "function",
            name: "exec",
            parameters: { type: "object" },
          },
        ],
        tool_choice: "auto",
      },
      options(),
    );
    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.tooling.toolChoiceMode).toBe(
      ToolChoiceModes.Auto,
    );
  });

  test("preserves mapped Responses instructions and input bytes", () => {
    const input = "  Keep this input exactly.  ";
    const instructions = "  Keep these instructions exactly.  ";
    const parsed = tryParseResponsesRequest(
      { model: "gpt-5.6-sol", input, instructions },
      mappedOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.base.promptText).toBe(input);
    expect(parsed.request.instructions).toBe(instructions);
  });
});

describe("mapped OpenAI text fidelity", () => {
  test("does not replace a user request with its embedded JSON object", () => {
    const content = [
      "Please preserve this formatted example:",
      "{",
      '  "enabled": true',
      "}",
      "Then explain it.",
    ].join("\n");
    const parsed = tryParseOpenAiRequest(
      {
        model: "gpt-5.6-sol",
        messages: [{ role: "user", content }],
      },
      mappedOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.promptText).toBe(content);
  });

  test("concatenates text parts without deleting or inventing separators", () => {
    const parsed = tryParseOpenAiRequest(
      {
        model: "gpt-5.6-sol",
        messages: [{
          role: "user",
          content: [
            { type: "text", text: "  Repository" },
            { type: "text", text: " reconnaissance " },
            { type: "text", text: "is reconciled.  " },
          ],
        }],
      },
      mappedOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    expect(parsed.request.promptText).toBe(
      "  Repository reconnaissance is reconciled.  ",
    );
  });
});

describe("Responses tool call extraction from assistant history", () => {
  test("extracts a custom_tool_call embedded in an assistant content JSON blob", () => {
    const assistantBlob = {
      id: "resp_custom_history",
      object: "response",
      status: "completed",
      model: "gpt-5.6-sol",
      output: [
        {
          type: "custom_tool_call",
          call_id: "call_abc",
          name: "exec",
          input: 'const r = await tools.exec_command({cmd:"ls"}); text(r);',
        },
      ],
    };

    const parsed = tryParseOpenAiRequest(
      {
        model: "gpt-5.6-sol",
        messages: [
          { role: "assistant", content: JSON.stringify(assistantBlob) },
          { role: "user", content: "What was the result?" },
        ],
      },
      mappedOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    const toolCalls = extractAssistantToolCallsContext(
      parsed.request.additionalContext,
    );
    expect(toolCalls).toHaveLength(1);
    expect(toolCalls[0].id).toBe("call_abc");
    expect(toolCalls[0].type).toBe("custom");
    expect(toolCalls[0].function.name).toBe("exec");
    expect(toolCalls[0].function.arguments).toBe(
      'const r = await tools.exec_command({cmd:"ls"}); text(r);',
    );
  });

  test("still extracts a function_call embedded in an assistant content JSON blob", () => {
    const assistantBlob = {
      id: "resp_function_history",
      object: "response",
      status: "completed",
      model: "gpt-5.6-sol",
      output: [
        {
          type: "function_call",
          call_id: "call_xyz",
          name: "exec",
          arguments: JSON.stringify({ cmd: "git status --short" }),
        },
      ],
    };

    const parsed = tryParseOpenAiRequest(
      {
        model: "gpt-5.6-sol",
        messages: [
          { role: "assistant", content: JSON.stringify(assistantBlob) },
          { role: "user", content: "Continue." },
        ],
      },
      mappedOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (!parsed.ok) return;
    const toolCalls = extractAssistantToolCallsContext(
      parsed.request.additionalContext,
    );
    expect(toolCalls).toHaveLength(1);
    expect(toolCalls[0].id).toBe("call_xyz");
    expect(toolCalls[0].type).toBe("function");
    expect(toolCalls[0].function.name).toBe("exec");
    expect(toolCalls[0].function.arguments).toBe(
      JSON.stringify({ cmd: "git status --short" }),
    );
  });
});

function extractAssistantToolCallsContext(
  additionalContext: { text: string; description: string | null }[],
): Array<{ id: string; type: string; function: { name: string; arguments: string } }> {
  const prefix = "assistant tool_calls: ";
  const entry = additionalContext.find((item) => item.text.includes(prefix));
  expect(entry).toBeDefined();
  const startIndex = (entry?.text ?? "").indexOf(prefix) + prefix.length;
  return JSON.parse((entry?.text ?? "").slice(startIndex));
}

function options(): WrapperOptions {
  return {
    host: "127.0.0.1", port: 4000, tokenPath: "/tmp/test-token", browser: "chromium", browserChannel: null, headless: true, loginTimeoutMs: 1, tokenTimeoutMs: 1, graphBaseUrl: "https://example.invalid", createConversationPath: "/conversations", chatPathTemplate: "/conversations/{conversationId}/chat", chatOverStreamPathTemplate: "/conversations/{conversationId}/stream", defaultModel: "gpt-5.6-sol", defaultSystemMessage: null, includeConversationIdInResponseBody: true, conversationTtlMinutes: 1, openAiTransformMode: OpenAiTransformModes.Simulated, transport: TransportNames.Substrate, logLevel: LogLevels.Error, logDirectory: "/tmp", logRequestBodies: false, logResponseBodies: false, retrySimulatedToollessResponses: true, substrate: { hubUrl: "ws://example.invalid", handshakeTimeoutMs: 1, turnTimeoutMs: 1, invocationTarget: "update", invocationType: 1 }, temporaryChat: true,
  };
}

function mappedOptions(): WrapperOptions {
  return { ...options(), openAiTransformMode: OpenAiTransformModes.Mapped };
}
