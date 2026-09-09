import { randomUUID } from "node:crypto";
import {
  decodeAndValidateSimulatedOutput,
  type SimulatedOutputAccepted,
  type SimulatedOutputRejected,
  type SimulatedOutputResult,
} from "./simulated-output";
import {
  ResponseFormatTypes,
  type JsonObject,
  type ParsedOpenAiRequest,
} from "./types";
import { isJsonObject, repairJsonStringLexemes } from "./utils";

const finalMarker = "M365_FINAL_V1";
const functionMarker = "M365_FUNCTION_TOOL_CALL_V1";
const customMarker = "M365_CUSTOM_TOOL_CALL_V1";

export type SimulatedActionOutputV1Result =
  | { kind: "not_v1" }
  | SimulatedOutputAccepted
  | SimulatedOutputRejected;

export function decodeSimulatedActionOutputV1(
  assistantText: string,
  request: ParsedOpenAiRequest,
): SimulatedActionOutputV1Result {
  if (assistantText.startsWith(finalMarker)) {
    return decodeFinalOutput(assistantText, request);
  }
  if (assistantText.startsWith(functionMarker)) {
    return decodeFunctionOutput(assistantText, request);
  }
  if (assistantText.startsWith(customMarker)) {
    return decodeCustomOutput(assistantText, request);
  }
  return { kind: "not_v1" };
}

function decodeFinalOutput(
  assistantText: string,
  request: ParsedOpenAiRequest,
): SimulatedActionOutputV1Result {
  const finalText = readMarkerBody(assistantText, finalMarker);
  if (finalText === null) {
    return rejected("invalid_envelope");
  }
  if (
    request.responseFormat?.type === ResponseFormatTypes.JsonObject ||
    request.responseFormat?.type === ResponseFormatTypes.JsonSchema
  ) {
    return rejected("invalid_message");
  }

  return validateEnvelope(request, [{
    type: "message",
    status: "completed",
    role: "assistant",
    content: [{ type: "output_text", text: finalText }],
  }]);
}

function decodeFunctionOutput(
  assistantText: string,
  request: ParsedOpenAiRequest,
): SimulatedActionOutputV1Result {
  const body = readMarkerBody(assistantText, functionMarker);
  if (body === null) {
    return rejected("invalid_envelope");
  }

  const separator = body.indexOf("\n");
  if (separator <= 0) {
    return rejected("invalid_envelope");
  }
  const name = body.slice(0, separator);
  const argumentsObject = parseFunctionArguments(body.slice(separator + 1));
  if (argumentsObject === null) {
    return rejected("malformed_arguments");
  }

  const argumentsJson = JSON.stringify(argumentsObject);
  const result = validateEnvelope(request, [{
    type: "function_call",
    status: "completed",
    call_id: createCallId(),
    name,
    arguments: argumentsJson,
  }]);
  return enforceExactToolName(result, name);
}

function decodeCustomOutput(
  assistantText: string,
  request: ParsedOpenAiRequest,
): SimulatedActionOutputV1Result {
  const body = readMarkerBody(assistantText, customMarker);
  if (body === null) {
    return rejected("invalid_envelope");
  }

  const separator = body.indexOf("\n");
  if (separator <= 0) {
    return rejected("invalid_envelope");
  }
  const name = body.slice(0, separator);
  const input = body.slice(separator + 1);
  const result = validateEnvelope(request, [{
    type: "custom_tool_call",
    status: "completed",
    call_id: createCallId(),
    name,
    input,
  }]);
  return enforceExactToolName(result, name);
}

function readMarkerBody(text: string, marker: string): string | null {
  if (!text.startsWith(marker) || text[marker.length] !== "\n") {
    return null;
  }
  return text.slice(marker.length + 1);
}

function parseFunctionArguments(rawArguments: string): JsonObject | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(rawArguments);
  } catch {
    const repaired = repairJsonStringLexemes(rawArguments);
    if (repaired === rawArguments) {
      return null;
    }
    try {
      parsed = JSON.parse(repaired);
    } catch {
      return null;
    }
  }
  return isJsonObject(parsed) ? parsed : null;
}

function validateEnvelope(
  request: ParsedOpenAiRequest,
  output: JsonObject[],
): SimulatedOutputResult {
  return decodeAndValidateSimulatedOutput(
    JSON.stringify({
      object: "response",
      status: "completed",
      output,
    }),
    "responses",
    request.tooling,
  );
}

function enforceExactToolName(
  result: SimulatedOutputResult,
  requestedName: string,
): SimulatedActionOutputV1Result {
  if (result.kind !== "accepted") {
    return result;
  }
  const call = result.items.find((item) => item.kind !== "message");
  return call && call.call.name !== requestedName
    ? rejected("tool_namespace_mismatch")
    : result;
}

function createCallId(): string {
  return `call_m365_${randomUUID().replaceAll("-", "")}`;
}

function rejected(
  reason: SimulatedOutputRejected["reason"],
): SimulatedOutputRejected {
  return { kind: "rejected", reason };
}
