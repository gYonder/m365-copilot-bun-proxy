import { describe, expect, test } from "bun:test";
import { resolveSubstrateTone } from "../src/proxy/clients";
import { ObservedSubstrateTones } from "../src/proxy/substrate-capabilities";

describe("resolveSubstrateTone", () => {
  test("maps the sole Codex model to the observed M365 reasoning selector", () => {
    expect(resolveSubstrateTone("gpt-5.6-sol")).toBe("Gpt_5_6_Reasoning");
  });

  test("defaults unknown and empty models to magic tone", () => {
    expect(resolveSubstrateTone("some-unknown-model")).toBe("magic");
    expect(resolveSubstrateTone("")).toBe("magic");
    expect(resolveSubstrateTone(null)).toBe("magic");
    expect(resolveSubstrateTone(undefined)).toBe("magic");
  });

  test("is case-insensitive", () => {
    expect(resolveSubstrateTone("GPT-5.6-SOL")).toBe("Gpt_5_6_Reasoning");
  });

  test("keeps the reviewed compatibility selectors explicit", () => {
    expect(resolveSubstrateTone("m365-copilot-quick")).toBe(
      ObservedSubstrateTones.Chat,
    );
    expect(resolveSubstrateTone("m365-copilot-reasoning")).toBe(
      ObservedSubstrateTones.Reasoning,
    );
    expect(resolveSubstrateTone("m365-copilot-gpt5.5-quick")).toBe(
      ObservedSubstrateTones.Gpt55Chat,
    );
  });
});
