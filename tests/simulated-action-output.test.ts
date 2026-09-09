import { describe, expect, test } from "bun:test";
import {
  decodeSimulatedActionOutputV1,
  type SimulatedActionOutputV1Result,
} from "../src/proxy/simulated-action-output";
import {
  ResponseFormatTypes,
  ToolChoiceModes,
  type OpenAiToolDefinition,
  type ParsedOpenAiRequest,
} from "../src/proxy/types";

describe("simulated action output protocol V1", () => {
  test("requires the marker at byte zero and preserves final text", () => {
    expectResult(
      decodeSimulatedActionOutputV1(
        "\nM365_FINAL_V1\nnot a protocol marker",
        request(),
      ),
      "not_v1",
    );

    const finalText = "  exact final text\nwith trailing spaces  ";
    const result = decodeSimulatedActionOutputV1(
      `M365_FINAL_V1\n${finalText}`,
      request(),
    );

    expectAccepted(result);
    expect(result.outputText).toBe(finalText);
    expect(result.items).toHaveLength(1);
    expect(result.items[0]?.kind).toBe("message");
  });

  test("rejects final text for JSON response formats and call-required policies", () => {
    const jsonResult = decodeSimulatedActionOutputV1(
      "M365_FINAL_V1\n{\"answer\":true}",
      request({
        responseFormat: {
          type: ResponseFormatTypes.JsonObject,
          name: null,
          jsonSchema: null,
        },
      }),
    );
    expectRejected(jsonResult, "invalid_message");

    for (const tooling of [
      { ...request().tooling, tools: [functionTool("exec")], toolChoiceMode: ToolChoiceModes.Required },
      { ...request().tooling, tools: [functionTool("exec")], requiredByLocalAction: true },
      {
        ...request().tooling,
        tools: [functionTool("exec")],
        toolChoiceMode: ToolChoiceModes.Function,
        toolChoiceFunctionName: "exec",
        toolChoiceToolType: "function" as const,
      },
    ]) {
      const result = decodeSimulatedActionOutputV1(
        "M365_FINAL_V1\nmust call",
        request({ tooling }),
      );
      expectRejected(result, "tool_choice_violation");
    }
  });

  test("accepts a schema-valid function call with a bridge-local call id", () => {
    const result = decodeSimulatedActionOutputV1(
      'M365_FUNCTION_TOOL_CALL_V1\nexec\n{"path":"src/main.ts"}',
      request({ tooling: functionTooling() }),
    );

    expectAccepted(result);
    const item = result.items[0];
    expect(item?.kind).toBe("function_call");
    if (item?.kind !== "function_call") return;
    expect(item.call.id).toMatch(/^call_m365_[0-9a-f-]+$/);
    expect(item.call.name).toBe("exec");
    expect(item.call.argumentsJson).toBe('{"path":"src/main.ts"}');
  });

  test("allows only deterministic lexical repair for function arguments", () => {
    const result = decodeSimulatedActionOutputV1(
      `M365_FUNCTION_TOOL_CALL_V1\nexec\n{"path":"line
file"}`,
      request({ tooling: functionTooling() }),
    );

    expectAccepted(result);
    const item = result.items[0];
    expect(item?.kind).toBe("function_call");
    if (item?.kind !== "function_call") return;
    expect(JSON.parse(item.call.argumentsJson)).toEqual({ path: "line\nfile" });
  });

  test("rejects non-object, truncated, prose, and multiple function arguments", () => {
    for (const argumentsText of [
      "[]",
      "true",
      '{"path":"missing"',
      '{"path":"ok"} prose',
      '{"path":"one"} {"path":"two"}',
    ]) {
      const result = decodeSimulatedActionOutputV1(
        `M365_FUNCTION_TOOL_CALL_V1\nexec\n${argumentsText}`,
        request({ tooling: functionTooling() }),
      );
      expectRejected(result, "malformed_arguments");
    }
  });

  test("uses the existing function schema validation", () => {
    const result = decodeSimulatedActionOutputV1(
      'M365_FUNCTION_TOOL_CALL_V1\nexec\n{"path":42}',
      request({ tooling: functionTooling() }),
    );

    expectRejected(result, "schema_violation");
  });

  test("preserves custom input byte-for-byte after the tool name", () => {
    const input = "*** arbitrary custom input \r\n  α\n";
    const result = decodeSimulatedActionOutputV1(
      `M365_CUSTOM_TOOL_CALL_V1\nrun_code\n${input}`,
      request({ tooling: customTooling("run_code") }),
    );

    expectAccepted(result);
    const item = result.items[0];
    expect(item?.kind).toBe("custom_tool_call");
    if (item?.kind !== "custom_tool_call") return;
    expect(item.call.argumentsJson).toBe(input);
  });

  test("rejects wrong, unknown, and non-exact tool names", () => {
    expectRejected(
      decodeSimulatedActionOutputV1(
        'M365_FUNCTION_TOOL_CALL_V1\nmissing\n{}',
        request({ tooling: functionTooling() }),
      ),
      "unknown_tool",
    );
    expectRejected(
      decodeSimulatedActionOutputV1(
        "M365_FUNCTION_TOOL_CALL_V1\napply_patch\n{}",
        request({ tooling: customTooling("apply_patch") }),
      ),
      "tool_type_mismatch",
    );
    expectRejected(
      decodeSimulatedActionOutputV1(
        'M365_CUSTOM_TOOL_CALL_V1\nexec\nraw',
        request({ tooling: functionTooling() }),
      ),
      "tool_type_mismatch",
    );
    expectRejected(
      decodeSimulatedActionOutputV1(
        'M365_FUNCTION_TOOL_CALL_V1\nfunctions.exec\n{"path":"x"}',
        request({ tooling: functionTooling() }),
      ),
      "tool_namespace_mismatch",
    );
  });

  test("recognizes malformed markers without falling back to legacy output", () => {
    expectRejected(
      decodeSimulatedActionOutputV1("M365_FINAL_V1", request()),
      "invalid_envelope",
    );
    expectRejected(
      decodeSimulatedActionOutputV1(
        "M365_FUNCTION_TOOL_CALL_V1\nexec",
        request({ tooling: functionTooling() }),
      ),
      "invalid_envelope",
    );
    expectRejected(
      decodeSimulatedActionOutputV1(
        "M365_CUSTOM_TOOL_CALL_V1\n",
        request({ tooling: customTooling("apply_patch") }),
      ),
      "invalid_envelope",
    );
  });
});

function request(
  overrides: Partial<ParsedOpenAiRequest> = {},
): ParsedOpenAiRequest {
  return {
    model: "m365-copilot",
    stream: false,
    transformMode: "simulated",
    promptText: "test",
    userKey: null,
    locationHint: {},
    contextualResources: null,
    additionalContext: [],
    tooling: {
      tools: [],
      toolChoiceMode: ToolChoiceModes.None,
      toolChoiceFunctionName: null,
      toolChoiceToolType: null,
      parallelToolCalls: false,
    },
    responseFormat: null,
    reasoningEffort: null,
    temperature: null,
    ...overrides,
  };
}

function functionTool(name: string): OpenAiToolDefinition {
  return {
    name,
    type: "function",
    description: null,
    parameters: {
      type: "object",
      properties: { path: { type: "string" } },
      required: ["path"],
      additionalProperties: false,
    },
    format: null,
  };
}

function functionTooling(): ParsedOpenAiRequest["tooling"] {
  return {
    tools: [functionTool("exec")],
    toolChoiceMode: ToolChoiceModes.Auto,
    toolChoiceFunctionName: null,
    toolChoiceToolType: null,
    parallelToolCalls: false,
  };
}

function customTooling(name: string): ParsedOpenAiRequest["tooling"] {
  return {
    tools: [{
      name,
      type: "custom",
      description: null,
      parameters: {},
      format: null,
    }],
    toolChoiceMode: ToolChoiceModes.Auto,
    toolChoiceFunctionName: null,
    toolChoiceToolType: null,
    parallelToolCalls: false,
  };
}

function expectAccepted(
  result: SimulatedActionOutputV1Result,
): asserts result is Extract<SimulatedActionOutputV1Result, { kind: "accepted" }> {
  expect(result.kind).toBe("accepted");
}

function expectRejected(
  result: SimulatedActionOutputV1Result,
  reason: Extract<SimulatedActionOutputV1Result, { kind: "rejected" }>["reason"],
): void {
  expect(result).toEqual({ kind: "rejected", reason });
}

function expectResult(
  result: SimulatedActionOutputV1Result,
  kind: SimulatedActionOutputV1Result["kind"],
): void {
  expect(result.kind).toBe(kind);
}
