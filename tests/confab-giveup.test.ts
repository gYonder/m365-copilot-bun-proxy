import { describe, expect, test } from "bun:test";
import { looksLikeConfabGiveUp } from "../src/proxy/server";

describe("looksLikeConfabGiveUp", () => {
  test("detects the canonical give-up phrasings", () => {
    const samples = [
      "It seems the file-editing tools are not available right now. Please restart the task.",
      "The file editing tools are currently not available.",
      "I don't have access to the file-editing tools needed for this.",
      "I'm unable to use the code tools, please try to restart the task.",
    ];
    for (const sample of samples) {
      expect(looksLikeConfabGiveUp(sample)).toBeTrue();
    }
  });

  test("does not flag a normal successful answer", () => {
    const samples = [
      "",
      "   ",
      "Here is the refactored function you asked for.",
      "I edited the file and updated the tests accordingly.",
      "The available tools worked and the task is complete.",
    ];
    for (const sample of samples) {
      expect(looksLikeConfabGiveUp(sample)).toBeFalse();
    }
  });
});
