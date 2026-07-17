import type { JsonObject, JsonValue } from "./types";

export const ContextEstimatorVersion = "codex-json-v1";

export type ContextEstimate = {
  inputTokens: number;
  outputTokens: number;
  serializedInputBytes: number;
};

export type PreviousContextUsage = {
  inputTokens: number;
  windowId: string | null;
};

export function resolveContextInputTokens(
  currentInputTokens: number,
  currentWindowId: string | null,
  previous: PreviousContextUsage | null,
): number {
  if (!previous) return currentInputTokens;
  if (
    currentWindowId &&
    previous.windowId &&
    currentWindowId !== previous.windowId
  ) {
    return currentInputTokens;
  }

  // previous_response_id normally carries only the new delta. If the current
  // request is already close to the previous live-context size, treat it as a
  // full-history/rebased request and avoid counting the transcript twice.
  if (currentInputTokens >= Math.floor(previous.inputTokens * 0.6)) {
    return Math.max(currentInputTokens, previous.inputTokens);
  }
  return previous.inputTokens + currentInputTokens;
}

export function estimateResponsesContext(
  request: JsonObject,
  output: JsonObject[] = [],
): ContextEstimate {
  const inputEnvelope = selectInputEnvelope(request);
  const serializedInput = JSON.stringify(inputEnvelope);
  const serializedOutput = JSON.stringify(output);
  return {
    inputTokens: estimateJsonTokens(serializedInput),
    outputTokens: estimateJsonTokens(serializedOutput),
    serializedInputBytes: new TextEncoder().encode(serializedInput).byteLength,
  };
}

export function estimateJsonTokens(serialized: string): number {
  if (!serialized) return 0;
  let ascii = 0;
  let nonAscii = 0;
  for (const character of serialized) {
    if (character.codePointAt(0)! <= 0x7f) ascii += 1;
    else nonAscii += 1;
  }
  // Conservative approximation for structured Codex payloads. JSON syntax,
  // identifiers, schemas, shell text, and source code tokenize more densely
  // than ordinary prose, while non-ASCII text can approach one token/character.
  return Math.max(1, Math.ceil(ascii / 3.2) + Math.ceil(nonAscii / 1.5));
}

function selectInputEnvelope(request: JsonObject): JsonObject {
  const envelope: JsonObject = {};
  for (const [key, value] of Object.entries(request)) {
    if (key === "stream" || key === "store" || key === "include") continue;
    envelope[key] = cloneJson(value);
  }
  return envelope;
}

function cloneJson(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(cloneJson);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, cloneJson(item)]),
    );
  }
  return value;
}
