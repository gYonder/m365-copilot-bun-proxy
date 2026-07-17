import { describe, expect, test } from "bun:test";
import {
  estimateResponsesContext,
  resolveContextInputTokens,
} from "../src/proxy/context-accounting";

describe("Responses context accounting", () => {
  test("counts instructions, developer tools, schemas, grammar, calls, and results", () => {
    const small = estimateResponsesContext({ input: [{ type: "message", role: "user", content: "hello" }] });
    const complete = estimateResponsesContext({
      instructions: "x".repeat(4_000),
      input: [
        {
          type: "additional_tools",
          role: "developer",
          tools: [{
            type: "custom",
            name: "apply_patch",
            description: "y".repeat(2_000),
            format: { type: "grammar", syntax: "lark", definition: "z".repeat(4_000) },
          }],
        },
        { type: "message", role: "user", content: "hello" },
        { type: "function_call", call_id: "call_1", name: "exec", arguments: "{\"cmd\":\"pwd\"}" },
        { type: "function_call_output", call_id: "call_1", output: "result".repeat(500) },
      ],
    });
    expect(complete.inputTokens).toBeGreaterThan(small.inputTokens + 3_000);
    expect(complete.serializedInputBytes).toBeGreaterThan(10_000);
  });

  test("ignores transport-only response flags", () => {
    const base = { model: "gpt-5.6-sol", input: "hello" };
    const first = estimateResponsesContext(base);
    const second = estimateResponsesContext({ ...base, stream: true, store: false, include: ["reasoning.encrypted_content"] });
    expect(second.inputTokens).toBe(first.inputTokens);
  });

  test("is conservative for code-heavy JSON", () => {
    const source = "const value = await tools.exec_command({ cmd: 'rg --files' });\n".repeat(1_000);
    const result = estimateResponsesContext({ input: source });
    expect(result.inputTokens).toBeGreaterThan(source.length / 4);
  });

  test("accumulates delta continuations in the same window", () => {
    expect(
      resolveContextInputTokens(5_000, "window-1", {
        inputTokens: 40_000,
        windowId: "window-1",
      }),
    ).toBe(45_000);
  });

  test("does not double count full-history continuations", () => {
    expect(
      resolveContextInputTokens(38_000, "window-1", {
        inputTokens: 40_000,
        windowId: "window-1",
      }),
    ).toBe(40_000);
  });

  test("resets accounting when Codex starts a compacted window", () => {
    expect(
      resolveContextInputTokens(8_000, "window-2", {
        inputTokens: 96_000,
        windowId: "window-1",
      }),
    ).toBe(8_000);
  });
});
