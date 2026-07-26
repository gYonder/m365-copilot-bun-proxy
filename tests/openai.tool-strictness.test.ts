import { describe, expect, test } from "bun:test";
import { buildAssistantResponse } from "../src/proxy/openai";
import {
  OpenAiTransformModes,
  ToolChoiceModes,
  type ParsedOpenAiRequest,
} from "../src/proxy/types";

describe("buildAssistantResponse strict tool behavior", () => {
  test("returns strict error when tool_choice is function and no valid tool call is present", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      "I cannot call tools for this request.",
    );

    expect(response.toolCalls.length).toBe(0);
    expect(response.content).toBeNull();
    expect(response.strictToolErrorMessage).toBeString();
    expect(response.strictToolErrorMessage).toContain("get_time");
  });

  test("returns strict error when tool_choice is required and no valid tool call is present", () => {
    const request = createRequest(ToolChoiceModes.Required, null);
    const response = buildAssistantResponse(
      request,
      "Still no JSON tool call payload here.",
    );

    expect(response.toolCalls.length).toBe(0);
    expect(response.content).toBeNull();
    expect(response.strictToolErrorMessage).toBeString();
    expect(response.strictToolErrorMessage).toContain("tool_calls");
  });

  test("extracts tool call and clears strict error when valid JSON is present", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      '{"tool_calls":[{"name":"get_time","arguments":{"zone":"UTC"}}]}',
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls.length).toBe(1);
    expect(response.toolCalls[0]?.name).toBe("get_time");
    expect(response.content).toBeNull();
  });

  test("extracts tool call when invalid placeholder JSON appears before valid payload", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      [
        "Output shape: {\"tool_calls\":[{\"name\":\"get_time\",\"arguments\":{...}}]}",
        "Here is the correct payload:",
        "{\"tool_calls\":[{\"name\":\"get_time\",\"arguments\":{\"zone\":\"UTC\"}}]}",
      ].join("\n"),
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls.length).toBe(1);
    expect(response.toolCalls[0]?.name).toBe("get_time");
    expect(response.toolCalls[0]?.argumentsJson).toBe("{\"zone\":\"UTC\"}");
    expect(response.content).toBeNull();
  });



  test("extracts M365-style tool call aliases into canonical JSON arguments", () => {
    const request = createRequest(ToolChoiceModes.Required, null);
    const response = buildAssistantResponse(
      request,
      JSON.stringify({
        tool_calls: [
          {
            recipient_name: "get_time",
            parameters: {
              zone: "UTC",
            },
          },
        ],
      }),
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls.length).toBe(1);
    expect(response.toolCalls[0]?.name).toBe("get_time");
    expect(response.toolCalls[0]?.argumentsJson).toContain("\"zone\"");
  });

  test("unwraps a single code field for Codex custom tools", () => {
    const request = createRequest(ToolChoiceModes.Required, null);
    request.tooling.tools = [{
      name: "apply_patch",
      type: "custom",
      description: "Apply a patch",
      parameters: {},
      format: null,
    }];
    const patch = "*** Begin Patch\n*** End Patch";
    const response = buildAssistantResponse(
      request,
      JSON.stringify({
        name: "apply_patch",
        input: { input: patch },
      }),
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls).toHaveLength(1);
    expect(response.toolCalls[0]?.type).toBe("custom");
    expect(response.toolCalls[0]?.argumentsJson).toBe(patch);
  });

  test("does not fabricate a tool call when strict local exec output is missing", () => {
    const request = createRequest(ToolChoiceModes.Required, null, {
      promptText:
        "You are responding through a local Codex harness with shell and file tools.\nUser request:\nRun pwd using the local shell tool.",
      toolName: "exec_command",
    });
    const response = buildAssistantResponse(
      request,
      "```text\n/mnt/file_upload\n```",
    );

    expect(response.toolCalls.length).toBe(0);
    expect(response.finishReason).toBe("stop");
    expect(response.strictToolErrorMessage).not.toBeNull();
  });

  test("rejects a function call missing required arguments", () => {
    const response = buildAssistantResponse(
      createRequest(ToolChoiceModes.Required, null),
      '{"tool_calls":[{"name":"get_time","arguments":{}}]}',
    );

    expect(response.toolCalls).toEqual([]);
    expect(response.strictToolErrorMessage).toContain("tool_calls");
  });

  test("rejects an unknown offered tool under strict choice", () => {
    const response = buildAssistantResponse(
      createRequest(ToolChoiceModes.Required, null),
      '{"tool_calls":[{"name":"not_offered","arguments":{"zone":"UTC"}}]}',
    );

    expect(response.toolCalls).toEqual([]);
    expect(response.strictToolErrorMessage).toContain("tool_calls");
  });
});

function createRequest(
  toolChoiceMode: string,
  toolChoiceFunctionName: string | null,
  overrides: {
    promptText?: string;
    toolName?: string;
  } = {},
): ParsedOpenAiRequest {
  const toolName = overrides.toolName ?? "get_time";
  return {
    model: "m365-copilot",
    stream: false,
    transformMode: OpenAiTransformModes.Mapped,
    promptText: overrides.promptText ?? "test",
    userKey: null,
    locationHint: { timeZone: "America/New_York" },
    contextualResources: null,
    additionalContext: [],
    tooling: {
      tools: [
        {
          name: toolName,
          description: "Get time",
          parameters: {
            type: "object",
            properties: { zone: { type: "string" } },
            required: ["zone"],
          },
        },
      ],
      toolChoiceMode,
      toolChoiceFunctionName,
      parallelToolCalls: true,
    },
    responseFormat: null,
    reasoningEffort: null,
    temperature: null,
  };
}
