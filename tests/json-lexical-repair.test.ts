import { describe, expect, test } from "bun:test";
import { repairJsonStringLexemes } from "../src/proxy/utils";

describe("repairJsonStringLexemes", () => {
  test("preserves valid JSON unchanged", () => {
    const raw = String.raw`{"pattern":"\\d+","line":"one\n two","quote":"\""}`;

    expect(repairJsonStringLexemes(raw)).toBe(raw);
  });

  test("repairs invalid escapes while preserving intended string bytes", () => {
    const raw = String.raw`{"regex":"\s+\d+\.","path":"C:\Users\workspace","unicode":"\u12G4"}`;
    const repaired = repairJsonStringLexemes(raw);

    expect(JSON.parse(repaired)).toEqual({
      regex: String.raw`\s+\d+\.`,
      path: String.raw`C:\Users\workspace`,
      unicode: String.raw`\u12G4`,
    });
  });

  test("repairs literal control characters only inside strings", () => {
    const raw = "{\"value\":\"line\nnext\"}";

    expect(JSON.parse(repairJsonStringLexemes(raw))).toEqual({
      value: "line\nnext",
    });
    expect(repairJsonStringLexemes("{\n\"value\":1}")).toBe(
      "{\n\"value\":1}",
    );
  });

  test("does not repair truncation or structural corruption", () => {
    const truncated = String.raw`{"value":"unfinished`;
    const structural = String.raw`{"value":"ok"} trailing`;

    expect(repairJsonStringLexemes(truncated)).toBe(truncated);
    expect(repairJsonStringLexemes(structural)).toBe(structural);
    expect(() => JSON.parse(repairJsonStringLexemes(truncated))).toThrow();
    expect(() => JSON.parse(repairJsonStringLexemes(structural))).toThrow();
  });
});
