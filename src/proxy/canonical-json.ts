import { createHash } from "node:crypto";

export const CANONICAL_JSON_MAX_DEPTH = 32;
export const CANONICAL_JSON_MAX_BYTES = 64 * 1024;
export const CANONICAL_JSON_MAX_NODES = 10_000;

export type CanonicalJsonErrorCode =
  | "invalid_options"
  | "max_depth_exceeded"
  | "max_size_exceeded"
  | "max_nodes_exceeded"
  | "non_finite_number"
  | "unsupported_value"
  | "cyclic_value";

export class CanonicalJsonError extends Error {
  readonly name = "CanonicalJsonError";

  constructor(
    readonly code: CanonicalJsonErrorCode,
    message: string,
    readonly path = "$",
  ) {
    super(message);
  }
}

export type CanonicalJsonResult<T = string> =
  | { ok: true; value: T }
  | { ok: false; error: CanonicalJsonError };

export type CanonicalJsonOptions = {
  maxDepth?: number;
  maxBytes?: number;
  maxNodes?: number;
};

const utf8 = new TextEncoder();

/**
 * Canonicalizes JSON-compatible values without changing array order.
 *
 * Object keys are sorted with JavaScript's relational string comparison,
 * which compares UTF-16 code units lexicographically. Numbers use
 * JavaScript's shortest round-trip representation; `-0` is emitted as `0`,
 * non-finite numbers are rejected, and JSON has no representation for them.
 * No Unicode normalization is performed, so distinct Unicode strings remain
 * distinct while their escaping is deterministic.
 */
export function canonicalizeJson(
  value: unknown,
  options: CanonicalJsonOptions = {},
): string {
  const result = tryCanonicalizeJson(value, options);
  if (!result.ok) {
    throw result.error;
  }
  return result.value;
}

export function tryCanonicalizeJson(
  value: unknown,
  options: CanonicalJsonOptions = {},
): CanonicalJsonResult {
  const limits = resolveLimits(options);
  if (!limits.ok) {
    return limits;
  }

  try {
    const output: string[] = [];
    const seen = new WeakSet<object>();
    let bytes = 0;
    let nodes = 0;

    const emit = (text: string): void => {
      bytes += utf8.encode(text).byteLength;
      if (bytes > limits.value.maxBytes) {
        throw new CanonicalJsonError(
          "max_size_exceeded",
          "Canonical JSON exceeds the configured size bound.",
        );
      }
      output.push(text);
    };

    const visit = (current: unknown, depth: number, path: string): void => {
      if (depth > limits.value.maxDepth) {
        throw new CanonicalJsonError(
          "max_depth_exceeded",
          "Canonical JSON exceeds the configured depth bound.",
          path,
        );
      }

      nodes += 1;
      if (nodes > limits.value.maxNodes) {
        throw new CanonicalJsonError(
          "max_nodes_exceeded",
          "Canonical JSON exceeds the configured node bound.",
          path,
        );
      }

      if (current === null) {
        emit("null");
        return;
      }

      switch (typeof current) {
        case "boolean":
          emit(current ? "true" : "false");
          return;
        case "string":
          emit(JSON.stringify(current));
          return;
        case "number":
          emit(formatNumber(current, path));
          return;
        case "object":
          break;
        default:
          throw new CanonicalJsonError(
            "unsupported_value",
            "Canonical JSON only accepts JSON-compatible values.",
            path,
          );
      }

      if (seen.has(current)) {
        throw new CanonicalJsonError(
          "cyclic_value",
          "Canonical JSON does not accept cyclic values.",
          path,
        );
      }
      if (!isPlainObjectOrArray(current)) {
        throw new CanonicalJsonError(
          "unsupported_value",
          "Canonical JSON only accepts plain objects and arrays.",
          path,
        );
      }

      seen.add(current);
      if (Array.isArray(current)) {
        emit("[");
        for (let index = 0; index < current.length; index += 1) {
          if (index > 0) emit(",");
          visit(current[index], depth + 1, `${path}[${index}]`);
        }
        emit("]");
      } else {
        const objectValue = current as Record<string, unknown>;
        const keys = Object.keys(objectValue).sort(compareJsonKeys);
        emit("{");
        for (let index = 0; index < keys.length; index += 1) {
          if (index > 0) emit(",");
          const key = keys[index];
          emit(JSON.stringify(key));
          emit(":");
          visit(objectValue[key], depth + 1, `${path}.${key}`);
        }
        emit("}");
      }
      seen.delete(current);
    };

    visit(value, 0, "$");
    return { ok: true, value: output.join("") };
  } catch (error) {
    if (error instanceof CanonicalJsonError) {
      return { ok: false, error };
    }
    return {
      ok: false,
      error: new CanonicalJsonError(
        "unsupported_value",
        "Canonical JSON could not be produced for this value.",
      ),
    };
  }
}

export function hashCanonicalJson(
  value: unknown,
  options: CanonicalJsonOptions = {},
): string {
  return createHash("sha256").update(canonicalizeJson(value, options)).digest("hex");
}

export function tryHashCanonicalJson(
  value: unknown,
  options: CanonicalJsonOptions = {},
): CanonicalJsonResult<string> {
  const result = tryCanonicalizeJson(value, options);
  if (!result.ok) {
    return result;
  }
  return {
    ok: true,
    value: createHash("sha256").update(result.value).digest("hex"),
  };
}

function formatNumber(value: number, path: string): string {
  if (!Number.isFinite(value)) {
    throw new CanonicalJsonError(
      "non_finite_number",
      "Canonical JSON rejects non-finite numbers.",
      path,
    );
  }
  if (Object.is(value, -0)) {
    return "0";
  }

  const text = value.toString();
  return text.replace(/e([+-]?)(\d+)$/i, (_match, sign: string, digits: string) => {
    const normalizedDigits = digits.replace(/^0+/, "") || "0";
    return `e${sign === "-" ? "-" : "+"}${normalizedDigits}`;
  });
}

function compareJsonKeys(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function isPlainObjectOrArray(value: object): boolean {
  if (Array.isArray(value)) {
    return true;
  }
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function resolveLimits(
  options: CanonicalJsonOptions,
): CanonicalJsonResult<{
  maxDepth: number;
  maxBytes: number;
  maxNodes: number;
}> {
  const maxDepth = options.maxDepth ?? CANONICAL_JSON_MAX_DEPTH;
  const maxBytes = options.maxBytes ?? CANONICAL_JSON_MAX_BYTES;
  const maxNodes = options.maxNodes ?? CANONICAL_JSON_MAX_NODES;

  if (
    !Number.isSafeInteger(maxDepth) ||
    maxDepth < 0 ||
    !Number.isSafeInteger(maxBytes) ||
    maxBytes <= 0 ||
    !Number.isSafeInteger(maxNodes) ||
    maxNodes <= 0
  ) {
    return {
      ok: false,
      error: new CanonicalJsonError(
        "invalid_options",
        "Canonical JSON bounds must be positive safe integers.",
      ),
    };
  }
  return { ok: true, value: { maxDepth, maxBytes, maxNodes } };
}
