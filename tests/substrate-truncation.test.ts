import { describe, expect, test } from "bun:test";
import { truncateSubstrateSendText } from "../src/proxy/clients";

describe("truncateSubstrateSendText", () => {
  test("returns the text unchanged when truncation is disabled", () => {
    const text = "x".repeat(1000);
    expect(truncateSubstrateSendText(text, 100, false)).toBe(text);
  });

  test("returns the text unchanged when under the limit", () => {
    const text = "short prompt";
    expect(truncateSubstrateSendText(text, 1000, true)).toBe(text);
  });

  test("preserves the trailing user prompt when truncating the middle", () => {
    const context = "CONTEXT-".repeat(2000);
    const tail = "\n\nUser: please answer THIS exact question.";
    const text = `Context:\n${context}${tail}`;
    const result = truncateSubstrateSendText(text, 1000, true);

    expect(result.length).toBeLessThanOrEqual(1000);
    // The user's actual prompt lives at the end and must survive truncation.
    expect(result.endsWith("please answer THIS exact question.")).toBeTrue();
    expect(result).toContain("characters truncated to fit Substrate limits");
    // The head is retained too, so the reader still sees the leading context.
    expect(result.startsWith("Context:")).toBeTrue();
  });

  test("keeps the tail when the marker cannot fit in a tiny budget", () => {
    const text = "abcdefghijklmnopqrstuvwxyz";
    const result = truncateSubstrateSendText(text, 5, true);
    expect(result).toBe("vwxyz");
  });
});
