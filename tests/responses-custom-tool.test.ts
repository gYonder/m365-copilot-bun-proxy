import { describe, expect, test } from "bun:test";
import { buildFunctionCallOutputItem } from "../src/proxy/responses-api";
import { tryParseResponsesRequest } from "../src/proxy/request-parser";
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
});

function options(): WrapperOptions {
  return {
    host: "127.0.0.1", port: 4000, tokenPath: "/tmp/test-token", browser: "chromium", browserChannel: null, headless: true, loginTimeoutMs: 1, tokenTimeoutMs: 1, graphBaseUrl: "https://example.invalid", createConversationPath: "/conversations", chatPathTemplate: "/conversations/{conversationId}/chat", chatOverStreamPathTemplate: "/conversations/{conversationId}/stream", defaultModel: "gpt-5.6-sol", defaultSystemMessage: null, includeConversationIdInResponseBody: true, conversationTtlMinutes: 1, openAiTransformMode: OpenAiTransformModes.Simulated, transport: TransportNames.Substrate, logLevel: LogLevels.Error, logDirectory: "/tmp", logRequestBodies: false, logResponseBodies: false, retrySimulatedToollessResponses: true, substrate: { hubUrl: "ws://example.invalid", handshakeTimeoutMs: 1, turnTimeoutMs: 1, invocationTarget: "update", invocationType: 1 }, temporaryChat: true,
  };
}
