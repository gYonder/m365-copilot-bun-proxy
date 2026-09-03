import { describe, expect, test } from "bun:test";
import {
  classifyToolAttempt,
  tryBuildAssistantResponseFromChatCompletionPayload,
  tryExtractIncrementalSimulatedChatContent,
  tryExtractSimulatedResponsePayload,
} from "../src/proxy/openai";
import { ToolChoiceModes, type OpenAiTooling } from "../src/proxy/types";

describe("tryExtractSimulatedResponsePayload", () => {
  test("extracts a markdown-wrapped Responses function call instead of exposing JSON as assistant text", () => {
    const responsePayload = {
      id: "resp_tool_regression",
      object: "response",
      status: "completed",
      model: "gpt-5.6-sol",
      output: [
        {
          type: "function_call",
          id: "fc_regression",
          call_id: "call_regression",
          name: "exec",
          arguments: JSON.stringify({
            cmd: "git status --short",
            workdir: "/workspace",
          }),
        },
      ],
    };
    const assistantText = [
      "```json",
      JSON.stringify(responsePayload, null, 2),
      "```",
    ].join("\n");

    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "responses",
    );

    expect(extracted).not.toBeNull();
    expect(extracted?.object).toBe("response");
    const output = Array.isArray(extracted?.output) ? extracted.output : [];
    expect(output).toHaveLength(1);
    expect(output[0]).toMatchObject({
      type: "function_call",
      call_id: "call_regression",
      name: "exec",
    });
  });

  test("extracts a markdown-wrapped Responses custom tool call without converting it to prose", () => {
    const responsePayload = {
      id: "resp_custom_tool_regression",
      object: "response",
      status: "completed",
      model: "gpt-5.6-sol",
      output: [
        {
          type: "custom_tool_call",
          id: "ctc_regression",
          call_id: "call_patch_regression",
          name: "apply_patch",
          input: "*** Begin Patch\n*** End Patch",
        },
      ],
    };
    const assistantText = [
      "```json",
      JSON.stringify(responsePayload, null, 2),
      "```",
    ].join("\n");

    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "responses",
    );

    expect(extracted).not.toBeNull();
    const output = Array.isArray(extracted?.output) ? extracted.output : [];
    expect(output).toHaveLength(1);
    expect(output[0]).toMatchObject({
      type: "custom_tool_call",
      call_id: "call_patch_regression",
      name: "apply_patch",
      input: "*** Begin Patch\n*** End Patch",
    });
  });

  test("does not accept a truncated markdown Responses tool call as a completed payload", () => {
    const truncated = [
      "```json",
      "{",
      "  \"object\": \"response\",",
      "  \"status\": \"completed\",",
      "  \"output\": [{",
      "    \"type\": \"function_call\",",
      "    \"name\": \"exec\",",
      "    \"arguments\": \"{\\\"cmd\\\":\\\"git status",
    ].join("\n");

    expect(
      tryExtractSimulatedResponsePayload(truncated, "responses"),
    ).toBeNull();
  });

  test("prefers chat completion response JSON over echoed request JSON", () => {
    const echoedRequest = {
      model: "gpt-5.1-2025-11-13",
      messages: [{ role: "user", content: "hello" }],
      tools: [
        {
          type: "function",
          function: { name: "write_to_file", parameters: { type: "object" } },
        },
      ],
      tool_choice: "auto",
      parallel_tool_calls: true,
    };

    const responsePayload = {
      id: "chatcmpl-test",
      object: "chat.completion",
      model: "gpt-5.1-2025-11-13",
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

    const assistantText = [
      "```json",
      JSON.stringify(echoedRequest, null, 2),
      "```",
      "```json",
      JSON.stringify(responsePayload, null, 2),
      "```",
    ].join("\n");

    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "chat.completions",
    );
    expect(extracted).not.toBeNull();
    expect(extracted?.id).toBe("chatcmpl-test");
    expect(Array.isArray(extracted?.choices)).toBeTrue();
  });

  test("returns null when only request-like JSON is present", () => {
    const requestOnly = {
      model: "gpt-5.1-2025-11-13",
      messages: [{ role: "user", content: "hello" }],
      tools: [
        {
          type: "function",
          function: { name: "write_to_file", parameters: { type: "object" } },
        },
      ],
      tool_choice: "required",
      parallel_tool_calls: true,
    };

    const assistantText = `\`\`\`json\n${JSON.stringify(requestOnly, null, 2)}\n\`\`\``;
    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "chat.completions",
    );
    expect(extracted).toBeNull();
  });
});

describe("tryExtractIncrementalSimulatedChatContent", () => {
  test("extracts partial content from incomplete simulated chat completion JSON", () => {
    const partial = [
      "```json",
      "{",
      '  "id": "chatcmpl_test",',
      '  "choices": [',
      "    {",
      '      "message": {',
      '        "role": "assistant",',
      '        "content": "hello\\nwor',
    ].join("\n");

    const extracted = tryExtractIncrementalSimulatedChatContent(partial);
    expect(extracted.hasToolCalls).toBeFalse();
    expect(extracted.content).toBe("hello\nwor");
  });

  test("suppresses incremental content when tool_calls appears before content", () => {
    const partial = [
      "```json",
      "{",
      '  "id": "chatcmpl_test",',
      '  "choices": [',
      "    {",
      '      "message": {',
      '        "role": "assistant",',
      '        "tool_calls": [',
      "          {",
      '            "id": "call_1",',
      '            "type": "function",',
      "            \"function\": {",
      '              "name": "attempt_completion",',
      "              \"arguments\": \"{\\\"result\\\":\\\"x\\\"}\"",
      "            }",
      "          }",
      "        ]",
    ].join("\n");

    const extracted = tryExtractIncrementalSimulatedChatContent(partial);
    expect(extracted.hasToolCalls).toBeTrue();
    expect(extracted.content).toBeNull();
  });
});

describe("simulated Chat Completions text fidelity", () => {
  test("preserves whitespace across assistant content-part boundaries", () => {
    const response = tryBuildAssistantResponseFromChatCompletionPayload({
      choices: [{
        finish_reason: "stop",
        message: {
          role: "assistant",
          content: [
            { type: "text", text: "Repository" },
            { type: "text", text: " reconnaissance " },
            { type: "text", text: "is reconciled." },
          ],
        },
      }],
    });

    expect(response?.content).toBe(
      "Repository reconnaissance is reconciled.",
    );
  });
});

describe("classifyToolAttempt", () => {
  const reportedEnvelope = JSON.stringify({
    type: "function_call",
    name: "functions.exec",
    call_id: "call_resume_catalog_implementation",
    arguments:
      'const r = await tools.exec_command({cmd:"sed -n \'80,340p\' src/observed-model-catalog.ts"}); text(r);',
  });

  function createTooling(toolName: string | null): OpenAiTooling {
    return {
      tools: toolName
        ? [
            {
              name: toolName,
              type: "custom",
              description: "Execute a tool",
              parameters: {},
              format: null,
            },
          ]
        : [],
      toolChoiceMode: ToolChoiceModes.Auto,
      toolChoiceFunctionName: null,
      parallelToolCalls: true,
    };
  }

  test("classifies functions.exec envelope as valid when exec tool is offered (prefix-strip)", () => {
    const classification = classifyToolAttempt(
      reportedEnvelope,
      createTooling("exec"),
    );
    expect(classification.kind).toBe("valid");
    if (classification.kind === "valid") {
      expect(classification.calls).toHaveLength(1);
      expect(classification.calls[0]?.name).toBe("exec");
      expect(classification.calls[0]?.type).toBe("custom");
      // JS expression preserved verbatim
      expect(classification.calls[0]?.argumentsJson).toContain(
        "tools.exec_command",
      );
    }
  });

  test("classifies functions.exec envelope as invalid_attempt when only exec_command is offered", () => {
    const classification = classifyToolAttempt(
      reportedEnvelope,
      createTooling("exec_command"),
    );
    expect(classification.kind).toBe("invalid_attempt");
    if (classification.kind === "invalid_attempt") {
      expect(classification.reason).toBe("tool_call_attempt_rejected");
    }
  });

  test("classifies as none when no tools are offered", () => {
    const classification = classifyToolAttempt(
      reportedEnvelope,
      createTooling(null),
    );
    expect(classification.kind).toBe("none");
  });

  test("classifies pure prose as none when tools are offered", () => {
    const classification = classifyToolAttempt(
      "I will now read the catalog file to understand the model layout.",
      createTooling("exec"),
    );
    expect(classification.kind).toBe("none");
  });

  test("classifies a malformed Responses tool envelope as an invalid attempt", () => {
    const malformedEnvelope = `{
      "object": "response",
      "output": [{
        "type": "function_call",
        "name": "exec",
        "arguments": "const result = await tools.exec_command({
          cmd: "pwd"
        });
        text(result);"
      }]
    }`;

    expect(
      tryExtractSimulatedResponsePayload(malformedEnvelope, "responses"),
    ).toBeNull();
    expect(
      classifyToolAttempt(malformedEnvelope, createTooling("exec")),
    ).toEqual({
      kind: "invalid_attempt",
      reason: "malformed_tool_call_envelope",
    });
  });

  test.each([
    {
      label: "fenced",
      text: "```json\n{\n  \"id\": \"resp_retry1\",\n  \"object\": \"response\",\n  \"status\": \"in_progress\",\n  \"model\": \"gpt-5.6-sol\",\n  \"output\": const applied = await tools.apply_patch(patch); text(applied);\n}\n```",
    },
    {
      label: "unfenced",
      text: "{\"id\":\"resp_retry2\",\"object\":\"response\",\"status\":\"in_progress\",\"model\":\"gpt-5.6-sol\",\"output\":const result = await tools.exec_command({cmd:\"pwd\"});text(JSON.stringify(result));}",
    },
  ])(
    "classifies a $label malformed Responses wrapper containing local tool code as an invalid attempt",
    ({ text }) => {
      expect(tryExtractSimulatedResponsePayload(text, "responses")).toBeNull();
      expect(classifyToolAttempt(text, createTooling("exec"))).toEqual({
        kind: "invalid_attempt",
        reason: "malformed_simulated_response_envelope",
      });
    },
  );

  test("classifies a matching tool call envelope as valid", () => {
    const validEnvelope = JSON.stringify({
      type: "function_call",
      name: "exec",
      call_id: "call_direct",
      arguments: 'const r = await tools.exec_command({cmd:"ls"}); text(r);',
    });
    const classification = classifyToolAttempt(
      validEnvelope,
      createTooling("exec"),
    );
    expect(classification.kind).toBe("valid");
    if (classification.kind === "valid") {
      expect(classification.calls).toHaveLength(1);
      expect(classification.calls[0]?.name).toBe("exec");
    }
  });
});
