import { describe, expect, test } from "bun:test";
import {
  decodeAndValidateSimulatedOutput,
  type SimulatedOutputResult,
} from "../src/proxy/simulated-output";
import {
  ToolChoiceModes,
  type JsonObject,
  type OpenAiTooling,
} from "../src/proxy/types";

describe("strict simulated output boundary", () => {
  test("projects a complete Chat envelope without trusting provider metadata", () => {
    const result = decodeAndValidateSimulatedOutput(
      JSON.stringify({
        object: "chat.completion",
        choices: [{
          index: 0,
          finish_reason: "stop",
          message: { role: "assistant", content: "hello" },
        }],
        id: "provider-id",
        model: "provider-model",
        created: 1,
        usage: { total_tokens: 999 },
      }),
      "chat.completions",
      tooling(),
    );

    expectAccepted(result);
    expect(result).toMatchObject({
      endpoint: "chat.completions",
      status: "completed",
      finishReason: "stop",
      outputText: "hello",
      items: [{
        kind: "message",
        content: "hello",
        refusal: null,
      }],
    });
    expect("payload" in result).toBeFalse();
  });

  test("accepts one JSON fence and rejects prose, multiple objects, or arrays", () => {
    const payload = chatPayload("hello");
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        fenced(payload),
        "chat.completions",
        tooling(),
      ),
    );
    expectRejected(
      `${JSON.stringify(payload)}\n${JSON.stringify(payload)}`,
      "chat.completions",
      tooling(),
      "malformed_json",
    );
    expectRejected(
      `prefix ${JSON.stringify(payload)}`,
      "chat.completions",
      tooling(),
      "malformed_json",
    );
    expectRejected(
      `[${JSON.stringify(payload)}]`,
      "chat.completions",
      tooling(),
      "json_object_required",
    );
  });

  test("repairs literal newlines inside outer Responses strings", () => {
    const result = decodeAndValidateSimulatedOutput(
      `{"object":"response","status":"completed","output":[{"type":"message","status":"completed","role":"assistant","content":[{"type":"output_text","text":"line one
line two"}]}],"output_text":"line one
line two"}`,
      "responses",
      tooling(),
    );

    expectAccepted(result);
    expect(result.outputText).toBe("line one\nline two");
  });

  test("repairs every literal JSON control character inside strings", () => {
    const controls = "\u0000\b\f\n\r\t\u001f";
    const valid = JSON.stringify(responsesPayload([responseMessage(controls)]));
    const malformed = valid
      .replaceAll("\\u0000", "\u0000")
      .replaceAll("\\b", "\b")
      .replaceAll("\\f", "\f")
      .replaceAll("\\n", "\n")
      .replaceAll("\\r", "\r")
      .replaceAll("\\t", "\t")
      .replaceAll("\\u001f", "\u001f");

    const result = decodeAndValidateSimulatedOutput(
      malformed,
      "responses",
      tooling(),
    );

    expectAccepted(result);
    expect(result.outputText).toBe(controls);
  });

  test("repairs literal newlines in nested function arguments", () => {
    const result = decodeAndValidateSimulatedOutput(
      `{"object":"response","status":"completed","output":[{"type":"function_call","status":"completed","call_id":"call_exec","name":"exec","arguments":"{\\"cmd\\":\\"line one
line two\\"}"}]}`,
      "responses",
      functionTooling(),
    );

    expectAccepted(result);
    const item = result.items[0];
    expect(item?.kind).toBe("function_call");
    if (item?.kind !== "function_call") return;
    expect(JSON.parse(item.call.argumentsJson)).toEqual({
      cmd: "line one\nline two",
    });
  });

  test("does not repair malformed JSON structure or controls outside strings", () => {
    const payload = JSON.stringify(responsesPayload([responseMessage("hello")]));
    expectRejected(
      payload.slice(0, -1),
      "responses",
      tooling(),
      "malformed_json",
    );
    expectRejected(
      `${payload}\n\u0000`,
      "responses",
      tooling(),
      "malformed_json",
    );
  });

  test("does not recursively inspect valid message text", () => {
    const content = [
      "Example JSON: ",
      JSON.stringify({ type: "function_call", name: "exec", arguments: "{}" }),
      "\n```ts\nconst x = \"*** Begin Patch\";\n```",
    ].join("");
    const result = decodeAndValidateSimulatedOutput(
      JSON.stringify(chatPayload(content)),
      "chat.completions",
      functionTooling(),
    );
    expectAccepted(result);
    expect(result.outputText).toBe(content);
    expect(result.items[0]?.kind).toBe("message");
  });

  test("requires endpoint marker and semantic envelope, not bridge metadata", () => {
    const missingMarker = chatPayload("hello");
    delete missingMarker.object;
    expectRejected(
      JSON.stringify(missingMarker),
      "chat.completions",
      tooling(),
      "missing_metadata",
    );
    expectRejected(
      JSON.stringify({
        object: "response",
        output: [],
        status: "completed",
      }),
      "responses",
      tooling(),
      "invalid_message",
    );
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        JSON.stringify({
          object: "chat.completion",
          choices: [{
            index: 0,
            finish_reason: "stop",
            message: { role: "assistant", content: "metadata-free" },
          }],
        }),
        "chat.completions",
        tooling(),
      ),
    );
  });

  test("rejects final in_progress and inconsistent Responses item status", () => {
    expectRejected(
      JSON.stringify(responsesPayload([responseMessage("hello")], "in_progress")),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      "unsupported_status",
    );
    const inconsistent = responsesPayload([responseMessage("hello")]);
    (inconsistent.output as JsonObject[])[0]!.status = "incomplete";
    expectRejected(
      JSON.stringify(inconsistent),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      "invalid_envelope",
    );
  });

  test("models terminal item status independently and never accepts failed items", () => {
    const incomplete = responsesPayload(
      [responseMessage("partial")],
      "incomplete",
    );
    (incomplete.output as JsonObject[])[0]!.status = "incomplete";
    incomplete.incomplete_details = { reason: "provider deadline" };
    const incompleteResult = decodeAndValidateSimulatedOutput(
      JSON.stringify(incomplete),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
    );
    expectAccepted(incompleteResult);
    expect(incompleteResult.items[0]).toMatchObject({
      kind: "message",
      status: "incomplete",
    });

    const failed = responsesPayload([responseMessage("partial")], "failed");
    failed.error = { code: "provider_failed", message: "provider failed" };
    const failedResult = decodeAndValidateSimulatedOutput(
      JSON.stringify(failed),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
    );
    expectAccepted(failedResult);
    expect(failedResult.items[0]?.status).toBe("completed");

    const failedItem = responsesPayload(
      [responseMessage("partial")],
      "failed",
    );
    (failedItem.output as JsonObject[])[0]!.status = "failed";
    failedItem.error = { code: "provider_failed", message: "provider failed" };
    expectRejected(
      JSON.stringify(failedItem),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      "invalid_envelope",
    );
  });

  test("allows omitted output_text but enforces exact text when present", () => {
    const omitted = responsesPayload([responseMessage("hello")]);
    delete omitted.output_text;
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        JSON.stringify(omitted),
        "responses",
        tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      ),
    );
    const mismatched = responsesPayload([responseMessage("hello")]);
    mismatched.output_text = "different";
    expectRejected(
      JSON.stringify(mismatched),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      "invalid_envelope",
    );
  });

  test("preserves mixed Responses ordering and exact call bytes", () => {
    const argumentsJson = '{ "cmd": "printf \\"x\\"" }';
    const input = "*** Begin Patch\r\n*** Add File: exact.txt\r\n+exact\r\n*** End Patch";
    const result = decodeAndValidateSimulatedOutput(
      JSON.stringify(responsesPayload([
        responseMessage("before"),
        responseFunctionCall("exec", argumentsJson, "call_middle"),
        responseCustomCall("apply_patch", input, "call_last"),
        responseMessage("after", "msg_after"),
      ])),
      "responses",
      customAndFunctionTooling(),
    );

    expectAccepted(result);
    expect(result.items.map((item) => item.kind)).toEqual([
      "message",
      "function_call",
      "custom_tool_call",
      "message",
    ]);
    expect(result.items[1]).toMatchObject({
      kind: "function_call",
      call: { id: "call_middle", argumentsJson },
    });
    expect(result.items[2]).toMatchObject({
      kind: "custom_tool_call",
      call: { id: "call_last", argumentsJson: input },
    });
  });

  test("resolves qualified names only through a compatible declared namespace", () => {
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        JSON.stringify(responsesPayload([
          responseFunctionCall("functions.exec", "{}", "call_fn"),
        ])),
        "responses",
        functionTooling(),
      ),
    );
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        JSON.stringify(responsesPayload([
          responseCustomCall("functions.exec", "exact input", "call_custom"),
        ])),
        "responses",
        tooling({
          tools: [{
            ...customTool("exec"),
            namespace: "functions",
          }],
          toolChoiceMode: ToolChoiceModes.Auto,
          parallelToolCalls: true,
        }),
      ),
    );
    expectRejected(
      JSON.stringify(responsesPayload([
        responseCustomCall("functions.apply_patch", "raw", "call_custom"),
      ])),
      "responses",
      customTooling(),
      "tool_namespace_mismatch",
    );
  });

  test("treats a null Chat refusal as absent", () => {
    const payload = chatPayload("valid final answer");
    const choices = payload.choices as JsonObject[];
    const message = choices[0]?.message;
    if (isJsonObjectLike(message)) {
      message.refusal = null;
    }

    const result = decodeAndValidateSimulatedOutput(
      JSON.stringify(payload),
      "chat.completions",
      tooling(),
    );

    expectAcceptedResult(result);
    expect(result.items).toMatchObject([
      {
        kind: "message",
        content: "valid final answer",
        refusal: null,
      },
    ]);
  });

  test("enforces none, auto, required, named, and parallel policies", () => {
    const call = [responseFunctionCall("exec", "{}", "call_1")];
    expectRejected(
      JSON.stringify(responsesPayload(call)),
      "responses",
      functionTooling({ toolChoiceMode: ToolChoiceModes.None }),
      "tool_choice_violation",
    );
    expectAcceptedResult(
      decodeAndValidateSimulatedOutput(
        JSON.stringify(responsesPayload([responseMessage("final")])),
        "responses",
        functionTooling({ toolChoiceMode: ToolChoiceModes.Auto }),
      ),
    );
    expectRejected(
      JSON.stringify(responsesPayload([responseMessage("final")])),
      "responses",
      functionTooling({ toolChoiceMode: ToolChoiceModes.Required }),
      "tool_choice_violation",
    );
    expectRejected(
      JSON.stringify(responsesPayload([
        responseFunctionCall("other", "{}", "call_named"),
      ])),
      "responses",
      functionTooling({
        tools: [functionTool("exec"), functionTool("other")],
        toolChoiceMode: ToolChoiceModes.Function,
        toolChoiceFunctionName: "exec",
      }),
      "tool_choice_violation",
    );
    expectRejected(
      JSON.stringify(responsesPayload([
        responseFunctionCall("exec", "{}", "call_1"),
        responseFunctionCall("exec", "{}", "call_2"),
      ])),
      "responses",
      functionTooling({ parallelToolCalls: false }),
      "parallel_calls_violation",
    );
  });

  test("preserves refusal parts for local rendering", () => {
    const payload = responsesPayload([{
      id: "provider-item",
      type: "message",
      status: "completed",
      role: "assistant",
      content: [{ type: "refusal", refusal: "I cannot help." }],
    }]);
    delete payload.output_text;
    const result = decodeAndValidateSimulatedOutput(
      JSON.stringify(payload),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Auto }),
    );
    expectAccepted(result);
    expect(result.items).toMatchObject([{
      kind: "message",
      content: "",
      refusal: "I cannot help.",
      parts: [{ kind: "refusal", text: "I cannot help." }],
    }]);
  });

  test("accepts explicit failed and incomplete terminal forms", () => {
    const failed = responsesPayload([], "failed");
    failed.id = "provider-id";
    failed.model = "provider-model";
    failed.created_at = 1;
    failed.usage = { total_tokens: 999 };
    failed.error = { code: "provider_failed", message: "provider failed" };
    failed.output_text = "";
    const failedResult = decodeAndValidateSimulatedOutput(
      JSON.stringify(failed),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Required }),
    );
    expectAcceptedResult(failedResult);
    expect(failedResult.terminalCause).toEqual({
      kind: "failed",
      code: "provider_failed",
      message: "provider failed",
    });
    const incomplete = responsesPayload([], "incomplete");
    incomplete.incomplete_details = { reason: " provider deadline " };
    incomplete.output_text = "";
    const incompleteResult = decodeAndValidateSimulatedOutput(
      JSON.stringify(incomplete),
      "responses",
      tooling({ toolChoiceMode: ToolChoiceModes.Required }),
    );
    expectAcceptedResult(incompleteResult);
    expect(incompleteResult.terminalCause).toEqual({
      kind: "incomplete",
      reason: "provider deadline",
    });
  });

  test("strictly validates Chat finish reasons and call consistency", () => {
    const message = {
      role: "assistant",
      content: null,
      tool_calls: [{
        id: "call_1",
        type: "function",
        function: { name: "exec", arguments: "{}" },
      }],
    };
    const callsWithStop = chatPayload("");
    callsWithStop.choices = [{
      index: 0,
      finish_reason: "stop",
      message,
    }];
    expectRejected(
      JSON.stringify(callsWithStop),
      "chat.completions",
      functionTooling(),
      "invalid_envelope",
    );

    const emptyToolBatch = chatPayload("");
    emptyToolBatch.choices = [{
      index: 0,
      finish_reason: "tool_calls",
      message: { role: "assistant", content: null, tool_calls: [] },
    }];
    expectRejected(
      JSON.stringify(emptyToolBatch),
      "chat.completions",
      functionTooling(),
      "invalid_envelope",
    );

    const missingToolBatch = chatPayload("");
    missingToolBatch.choices = [{
      index: 0,
      finish_reason: "function_call",
      message: { role: "assistant", content: null },
    }];
    expectRejected(
      JSON.stringify(missingToolBatch),
      "chat.completions",
      functionTooling(),
      "invalid_envelope",
    );

    const unsupportedFinish = chatPayload("done");
    unsupportedFinish.choices = [{
      index: 0,
      finish_reason: "provider_done",
      message: { role: "assistant", content: "done" },
    }];
    expectRejected(
      JSON.stringify(unsupportedFinish),
      "chat.completions",
      tooling(),
      "invalid_envelope",
    );

    const legacyFunctionCall = chatPayload("");
    legacyFunctionCall.choices = [{
      index: 0,
      finish_reason: "function_call",
      message: {
        role: "assistant",
        content: null,
        function_call: { name: "exec", arguments: "{}" },
      },
    }];
    const legacyResult = decodeAndValidateSimulatedOutput(
      JSON.stringify(legacyFunctionCall),
      "chat.completions",
      functionTooling(),
    );
    expectAccepted(legacyResult);
    expect(legacyResult.items[0]?.kind).toBe("function_call");
    expect(legacyResult.finishReason).toBe("tool_calls");
  });

  test("rejects duplicate call IDs before a response can be projected", () => {
    const duplicateChat = chatPayload("");
    duplicateChat.choices = [{
      index: 0,
      finish_reason: "tool_calls",
      message: {
        role: "assistant",
        content: null,
        tool_calls: [
          {
            id: "call_same",
            type: "function",
            function: { name: "exec", arguments: "{}" },
          },
          {
            id: "call_same",
            type: "function",
            function: { name: "exec", arguments: "{}" },
          },
        ],
      },
    }];
    expectRejected(
      JSON.stringify(duplicateChat),
      "chat.completions",
      functionTooling(),
      "duplicate_call_id",
    );

    const duplicateResponses = responsesPayload([
      responseFunctionCall("exec", "{}", "call_same"),
      responseFunctionCall("exec", "{}", "call_same"),
    ]);
    expectRejected(
      JSON.stringify(duplicateResponses),
      "responses",
      functionTooling(),
      "duplicate_call_id",
    );
  });
});

function chatPayload(content: string): JsonObject {
  return {
    object: "chat.completion",
    choices: [{
      index: 0,
      finish_reason: "stop",
      message: { role: "assistant", content },
    }],
  };
}

function responsesPayload(
  output: JsonObject[],
  status = "completed",
): JsonObject {
  return {
    object: "response",
    status,
    output,
    output_text: output
      .filter((item) => item.type === "message")
      .flatMap((item) =>
        Array.isArray(item.content)
          ? item.content
              .filter((part): part is JsonObject =>
                isJsonObjectLike(part) &&
                part.type === "output_text" &&
                typeof part.text === "string"
              )
              .map((part) => part.text as string)
          : []
      )
      .join(""),
  };
}

function responseMessage(content: string, id = "msg_provider"): JsonObject {
  return {
    id,
    type: "message",
    status: "completed",
    role: "assistant",
    content: [{ type: "output_text", text: content, annotations: [] }],
  };
}

function responseFunctionCall(
  name: string,
  argumentsJson: string,
  callId = "call_provider",
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

function responseCustomCall(
  name: string,
  input: string,
  callId = "call_provider",
): JsonObject {
  return {
    id: `item_${callId}`,
    type: "custom_tool_call",
    status: "completed",
    call_id: callId,
    name,
    input,
  };
}

function fenced(payload: JsonObject): string {
  return `\n\`\`\`json\n${JSON.stringify(payload, null, 2)}\n\`\`\`\n`;
}

function tooling(overrides: Partial<OpenAiTooling> = {}): OpenAiTooling {
  return {
    tools: [],
    toolChoiceMode: ToolChoiceModes.None,
    toolChoiceFunctionName: null,
    parallelToolCalls: false,
    ...overrides,
  };
}

function functionTool(name: string): OpenAiTooling["tools"][number] {
  return {
    name,
    type: "function",
    description: null,
    parameters: { type: "object" },
    format: null,
  };
}

function customTool(name: string): OpenAiTooling["tools"][number] {
  return {
    name,
    type: "custom",
    description: null,
    parameters: {},
    format: null,
  };
}

function functionTooling(
  overrides: Partial<OpenAiTooling> = {},
): OpenAiTooling {
  return tooling({
    tools: [functionTool("exec")],
    toolChoiceMode: ToolChoiceModes.Auto,
    parallelToolCalls: true,
    ...overrides,
  });
}

function customTooling(
  overrides: Partial<OpenAiTooling> = {},
): OpenAiTooling {
  return tooling({
    tools: [customTool("apply_patch")],
    toolChoiceMode: ToolChoiceModes.Auto,
    parallelToolCalls: true,
    ...overrides,
  });
}

function customAndFunctionTooling(): OpenAiTooling {
  return tooling({
    tools: [functionTool("exec"), customTool("apply_patch")],
    toolChoiceMode: ToolChoiceModes.Auto,
    parallelToolCalls: true,
  });
}

function isJsonObjectLike(value: unknown): value is JsonObject {
  return typeof value === "object" &&
    value !== null &&
    !Array.isArray(value);
}

function expectAcceptedResult(
  result: SimulatedOutputResult,
): asserts result is Extract<SimulatedOutputResult, { kind: "accepted" }> {
  expect(result.kind).toBe("accepted");
}

function expectAccepted(
  result: SimulatedOutputResult,
): asserts result is Extract<SimulatedOutputResult, { kind: "accepted" }> {
  expect(result.kind).toBe("accepted");
}

function expectRejected(
  text: string,
  endpoint: "chat.completions" | "responses",
  toolingValue: OpenAiTooling,
  reason: Extract<SimulatedOutputResult, { kind: "rejected" }>["reason"],
): void {
  expect(
    decodeAndValidateSimulatedOutput(text, endpoint, toolingValue),
  ).toEqual({ kind: "rejected", reason });
}
