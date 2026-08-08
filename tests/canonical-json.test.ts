import { describe, expect, test } from "bun:test";
import {
  canonicalizeJson,
  CanonicalJsonError,
  hashCanonicalJson,
  tryCanonicalizeJson,
} from "../src/proxy/canonical-json";

describe("canonical JSON", () => {
  test("orders object keys recursively", () => {
    expect(canonicalizeJson({ b: 2, a: { d: 4, c: 3 } })).toBe(
      '{"a":{"c":3,"d":4},"b":2}',
    );
    expect(canonicalizeJson({ a: 1, b: 2 })).toBe(
      canonicalizeJson({ b: 2, a: 1 }),
    );
  });

  test("preserves array order", () => {
    expect(canonicalizeJson({ values: [2, 1, { b: 2, a: 1 }] })).toBe(
      '{"values":[2,1,{"a":1,"b":2}]}',
    );
    expect(canonicalizeJson([1, 2])).not.toBe(canonicalizeJson([2, 1]));
  });

  test("normalizes whitespace and equivalent formatting", () => {
    const first = JSON.parse(' { "b": [ 1, 2 ], "a": true } ');
    const second = JSON.parse('{"a":true,"b":[1,2]}');
    expect(canonicalizeJson(first)).toBe(canonicalizeJson(second));
    expect(canonicalizeJson(first)).toBe('{"a":true,"b":[1,2]}');
  });

  test("uses stable numeric formatting", () => {
    expect(
      canonicalizeJson({
        integer: 2,
        float: 1.5,
        negativeZero: -0,
        large: 1e21,
        tiny: 1e-7,
      }),
    ).toBe(
      '{"float":1.5,"integer":2,"large":1e+21,"negativeZero":0,"tiny":1e-7}',
    );
    expect(tryCanonicalizeJson({ value: Number.NaN }).error.code).toBe(
      "non_finite_number",
    );
    expect(tryCanonicalizeJson({ value: Number.POSITIVE_INFINITY }).ok).toBe(
      false,
    );
  });

  test("keeps Unicode deterministic without normalizing it", () => {
    const value = { "\u00e5": "東京", "\ud83d\ude00": "cafe\u0301" };
    expect(canonicalizeJson(value)).toBe(
      '{"å":"東京","😀":"café"}',
    );
    expect(hashCanonicalJson(value)).toBe(hashCanonicalJson({ ...value }));
  });

  test("returns typed errors for cycles, depth, and size bounds", () => {
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;
    const cycleResult = tryCanonicalizeJson(cyclic);
    expect(cycleResult.ok).toBe(false);
    if (!cycleResult.ok) {
      expect(cycleResult.error).toBeInstanceOf(CanonicalJsonError);
      expect(cycleResult.error.code).toBe("cyclic_value");
    }

    let nested: unknown = null;
    for (let index = 0; index < 10; index += 1) {
      nested = { next: nested };
    }
    expect(tryCanonicalizeJson(nested, { maxDepth: 3 }).error.code).toBe(
      "max_depth_exceeded",
    );
    expect(
      tryCanonicalizeJson({ text: "0123456789" }, { maxBytes: 5 }).error.code,
    ).toBe("max_size_exceeded");
  });
});
