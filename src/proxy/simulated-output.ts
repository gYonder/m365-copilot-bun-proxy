import { randomUUID } from "node:crypto";
import {
  ToolChoiceModes,
  type JsonObject,
  type JsonValue,
  type OpenAiAssistantToolCall,
  type OpenAiTooling,
  type ParsedResponsesRequest,
} from "./types";
import {
  validateOpenAiToolCall,
  type OpenAiToolCallValidation,
} from "./openai";
import {
  buildFunctionCallOutputItem,
  buildMessageOutputItem,
  buildOpenAiResponseObject,
  createOpenAiOutputItemId,
  createOpenAiResponseId,
  type ResponseTerminal,
} from "./responses-api";
import { estimateJsonTokens } from "./context-accounting";
import { stripPrivateCitationMarkers } from "./responses-provenance";
import {
  escapeJsonControlCharactersInStrings,
  isJsonObject,
  nowUnix,
} from "./utils";

export type SimulatedOutputEndpoint = "chat.completions" | "responses";

export type SimulatedOutputRejectReason =
  | "empty_input"
  | "not_single_json_envelope"
  | "malformed_json"
  | "json_object_required"
  | "wrong_endpoint"
  | "missing_metadata"
  | "invalid_envelope"
  | "unsupported_status"
  | "invalid_message"
  | "unknown_tool"
  | "tool_namespace_mismatch"
  | "tool_type_mismatch"
  | "malformed_arguments"
  | "schema_violation"
  | "invalid_custom_input"
  | "tool_choice_violation"
  | "parallel_calls_violation"
  | "duplicate_call_id";

export type SimulatedOutputItemStatus =
  | "in_progress"
  | "completed"
  | "incomplete";

export type SimulatedOutputMessagePart =
  | { kind: "output_text"; text: string }
  | { kind: "refusal"; text: string };

export type SimulatedOutputItem =
  | {
      kind: "message";
      status: SimulatedOutputItemStatus;
      content: string;
      refusal: string | null;
      parts: SimulatedOutputMessagePart[];
    }
  | {
      kind: "function_call";
      status: SimulatedOutputItemStatus;
      call: OpenAiAssistantToolCall;
    }
  | {
      kind: "custom_tool_call";
      status: SimulatedOutputItemStatus;
      call: OpenAiAssistantToolCall;
    };

export type SimulatedOutputAccepted = {
  kind: "accepted";
  endpoint: SimulatedOutputEndpoint;
  status: "completed" | "failed" | "incomplete";
  finishReason: string | null;
  outputText: string;
  items: SimulatedOutputItem[];
  terminalCause: SimulatedOutputTerminalCause | null;
};

export type SimulatedOutputTerminalCause =
  | {
      kind: "failed";
      code: string;
      message: string;
    }
  | {
      kind: "incomplete";
      reason: string;
    };

export type SimulatedOutputRejected = {
  kind: "rejected";
  reason: SimulatedOutputRejectReason;
};

export type SimulatedOutputResult =
  | SimulatedOutputAccepted
  | SimulatedOutputRejected;

type SimulatedResponseStatus = SimulatedOutputAccepted["status"];

const functionNamespacePrefix = "functions.";

export function decodeAndValidateSimulatedOutput(
  assistantText: string,
  endpoint: SimulatedOutputEndpoint,
  tooling: OpenAiTooling,
): SimulatedOutputResult {
  const decoded = decodeJsonObject(assistantText);
  if (decoded.kind === "rejected") {
    return decoded;
  }

  return endpoint === "chat.completions"
    ? validateChatCompletion(decoded.payload, tooling)
    : validateResponses(decoded.payload, tooling);
}

function decodeJsonObject(
  assistantText: string,
): { kind: "accepted"; payload: JsonObject } | SimulatedOutputRejected {
  const trimmed = assistantText.trim();
  if (!trimmed) {
    return { kind: "rejected", reason: "empty_input" };
  }

  const jsonText = unwrapSingleJsonFence(trimmed);
  if (jsonText === null) {
    return { kind: "rejected", reason: "not_single_json_envelope" };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonText);
  } catch {
    const repaired = escapeJsonControlCharactersInStrings(jsonText);
    if (repaired === jsonText) {
      return { kind: "rejected", reason: "malformed_json" };
    }
    try {
      parsed = JSON.parse(repaired);
    } catch {
      return { kind: "rejected", reason: "malformed_json" };
    }
  }

  if (!isJsonObject(parsed)) {
    return { kind: "rejected", reason: "json_object_required" };
  }
  return { kind: "accepted", payload: parsed };
}

function unwrapSingleJsonFence(trimmed: string): string | null {
  if (!trimmed.startsWith("```")) {
    return trimmed;
  }

  const match = /^```json[ \t]*\r?\n([\s\S]*?)\r?\n```$/.exec(trimmed);
  const body = match?.[1];
  return body === undefined ||
      /(?:^|\r?\n)[ \t]*```[ \t]*(?:\r?\n|$)/.test(body)
    ? null
    : body;
}

function validateChatCompletion(
  payload: JsonObject,
  tooling: OpenAiTooling,
): SimulatedOutputResult {
  if (payload.object === undefined) {
    return { kind: "rejected", reason: "missing_metadata" };
  }
  if (payload.object !== "chat.completion") {
    return { kind: "rejected", reason: "wrong_endpoint" };
  }

  const choices = payload.choices;
  if (!Array.isArray(choices) || choices.length !== 1) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const choice = choices[0];
  if (!isJsonObject(choice) || !isCompleteChatChoice(choice)) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  const message = choice.message;
  if (!isJsonObject(message) || message.role !== "assistant") {
    return { kind: "rejected", reason: "invalid_message" };
  }

  const callsNode = message.tool_calls;
  if (callsNode !== undefined && !Array.isArray(callsNode)) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const legacyFunctionCall = message.function_call;
  if (
    legacyFunctionCall !== undefined &&
    (!isJsonObject(legacyFunctionCall) || callsNode !== undefined)
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const calls = callsNode
    ? validateChatCalls(callsNode, choice.finish_reason, tooling)
    : legacyFunctionCall !== undefined
      ? validateLegacyChatCall(legacyFunctionCall, tooling)
      : { kind: "accepted" as const, calls: [] };
  if (calls.kind === "rejected") {
    return calls;
  }
  if (
    calls.calls.length > 0 &&
    choice.finish_reason !== "tool_calls" &&
    choice.finish_reason !== "function_call"
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  if (
    calls.calls.length === 0 &&
    (choice.finish_reason === "tool_calls" ||
      choice.finish_reason === "function_call")
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  const messageResult = readChatMessage(message);
  if (messageResult.kind === "rejected") {
    return messageResult;
  }
  if (
    calls.calls.length === 0 &&
    messageResult.item === null
  ) {
    return { kind: "rejected", reason: "invalid_message" };
  }
  if (calls.calls.length > 0 && messageResult.refusal?.trim()) {
    return { kind: "rejected", reason: "invalid_message" };
  }
  const items = [
    ...(messageResult.item ? [messageResult.item] : []),
    ...calls.calls,
  ];
  const policy = validateToolPolicy(
    calls.calls,
    tooling,
    messageResult.refusal,
    "completed",
  );
  if (policy) {
    return policy;
  }

  return {
    kind: "accepted",
    endpoint: "chat.completions",
    status: "completed",
    finishReason:
      calls.calls.length > 0 && choice.finish_reason === "function_call"
        ? "tool_calls"
        : String(choice.finish_reason),
    outputText:
      messageResult.item?.kind === "message"
        ? messageResult.item.content
        : "",
    items,
    terminalCause: null,
  };
}

function isCompleteChatChoice(choice: JsonObject): boolean {
  return (
    typeof choice.index === "number" &&
    Number.isInteger(choice.index) &&
    choice.index >= 0 &&
    isSupportedChatFinishReason(choice.finish_reason) &&
    isJsonObject(choice.message)
  );
}

function isSupportedChatFinishReason(
  value: JsonValue | undefined,
): value is "stop" | "length" | "content_filter" | "tool_calls" | "function_call" {
  return (
    value === "stop" ||
    value === "length" ||
    value === "content_filter" ||
    value === "tool_calls" ||
    value === "function_call"
  );
}

function validateChatCalls(
  callsNode: JsonValue[],
  finishReason: JsonValue,
  tooling: OpenAiTooling,
): { kind: "accepted"; calls: SimulatedOutputItem[] } | SimulatedOutputRejected {
  if (callsNode.length === 0) {
    if (finishReason === "tool_calls" || finishReason === "function_call") {
      return { kind: "rejected", reason: "invalid_envelope" };
    }
    return { kind: "accepted", calls: [] };
  }
  if (finishReason !== "tool_calls" && finishReason !== "function_call") {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  const calls: SimulatedOutputItem[] = [];
  const callIds = new Set<string>();
  for (const node of callsNode) {
    const call = validateChatCall(node, tooling);
    if (call.kind === "rejected") {
      return call;
    }
    if (callIds.has(call.call.id)) {
      return { kind: "rejected", reason: "duplicate_call_id" };
    }
    callIds.add(call.call.id);
    calls.push({ kind: "function_call", status: "completed", call: call.call });
  }
  return { kind: "accepted", calls };
}

function validateLegacyChatCall(
  node: JsonObject,
  tooling: OpenAiTooling,
): { kind: "accepted"; calls: SimulatedOutputItem[] } | SimulatedOutputRejected {
  if (
    !isNonEmptyString(node.name) ||
    typeof node.arguments !== "string"
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const result = validateFunctionCall(
    isNonEmptyString(node.id)
      ? node.id
      : `call_${randomUUID().replaceAll("-", "")}`,
    node.name,
    node.arguments,
    tooling,
  );
  if (result.kind === "rejected") {
    return result;
  }
  return {
    kind: "accepted",
    calls: [{ kind: "function_call", status: "completed", call: result.call }],
  };
}

function validateChatCall(
  node: JsonValue,
  tooling: OpenAiTooling,
): { kind: "accepted"; call: OpenAiAssistantToolCall } | SimulatedOutputRejected {
  if (!isJsonObject(node) || node.type !== "function") {
    return { kind: "rejected", reason: "tool_type_mismatch" };
  }
  if (!isNonEmptyString(node.id) || !isJsonObject(node.function)) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const functionNode = node.function;
  if (
    !isNonEmptyString(functionNode.name) ||
    typeof functionNode.arguments !== "string"
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  return validateFunctionCall(
    node.id,
    functionNode.name,
    functionNode.arguments,
    tooling,
  );
}

function validateResponses(
  payload: JsonObject,
  tooling: OpenAiTooling,
): SimulatedOutputResult {
  if (payload.object === undefined) {
    return { kind: "rejected", reason: "missing_metadata" };
  }
  if (payload.object !== "response") {
    return { kind: "rejected", reason: "wrong_endpoint" };
  }

  const status = payload.status;
  if (status === undefined) {
    return { kind: "rejected", reason: "missing_metadata" };
  }
  if (!isResponseStatus(status)) {
    return { kind: "rejected", reason: "unsupported_status" };
  }

  const output = payload.output;
  if (!Array.isArray(output)) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const items: SimulatedOutputItem[] = [];
  const callIds = new Set<string>();
  for (const node of output) {
    const item = validateResponsesItem(node, status, tooling);
    if (item.kind === "rejected") {
      return item;
    }
    if (isCallItem(item.item)) {
      if (callIds.has(item.item.call.id)) {
        return { kind: "rejected", reason: "duplicate_call_id" };
      }
      callIds.add(item.item.call.id);
    }
    items.push(item.item);
  }

  const outputText = payload.output_text;
  if (outputText !== undefined &&
      (typeof outputText !== "string" || outputText !== responseText(items))) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  const terminal = validateResponsesTerminal(payload, status, items);
  if (terminal?.kind === "rejected") {
    return terminal;
  }

  const calls = items.filter(isCallItem);
  const message = items.find((item) => item.kind === "message");
  const policy = validateToolPolicy(
    calls,
    tooling,
    message?.kind === "message" ? message.refusal : null,
    status,
  );
  if (policy) {
    return policy;
  }

  return {
    kind: "accepted",
    endpoint: "responses",
    status,
    finishReason: null,
    outputText: responseText(items),
    items,
    terminalCause: terminal === null ? null : terminal,
  };
}

function validateResponsesTerminal(
  payload: JsonObject,
  status: SimulatedResponseStatus,
  items: SimulatedOutputItem[],
): SimulatedOutputRejected | SimulatedOutputTerminalCause | null {
  const error = payload.error;
  const incompleteDetails = payload.incomplete_details;

  if (status === "completed") {
    if (
      (error !== undefined && error !== null) ||
      (incompleteDetails !== undefined && incompleteDetails !== null) ||
      items.length === 0
    ) {
      return items.length === 0
        ? { kind: "rejected", reason: "invalid_message" }
        : { kind: "rejected", reason: "invalid_envelope" };
    }
    return null;
  }

  if (status === "failed") {
    if (
      !isJsonObject(error) ||
      !isNonEmptyString(error.code) ||
      !isNonEmptyString(error.message) ||
      (incompleteDetails !== undefined && incompleteDetails !== null)
    ) {
      return { kind: "rejected", reason: "invalid_envelope" };
    }
    return {
      kind: "failed",
      code: sanitizeTerminalField(error.code),
      message: sanitizeTerminalField(error.message),
    };
  } else if (
    !isJsonObject(incompleteDetails) ||
    !isNonEmptyString(incompleteDetails.reason) ||
    (error !== undefined && error !== null)
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }

  return {
    kind: "incomplete",
    reason: sanitizeTerminalField(incompleteDetails.reason),
  };
}

function sanitizeTerminalField(value: string): string {
  return value.trim().slice(0, 512);
}

function validateResponsesItem(
  node: JsonValue,
  outerStatus: SimulatedResponseStatus,
  tooling: OpenAiTooling,
): { kind: "accepted"; item: SimulatedOutputItem } | SimulatedOutputRejected {
  if (!isJsonObject(node)) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const itemStatus = node.status;
  if (
    !isItemStatus(itemStatus) ||
    (outerStatus === "completed"
      ? itemStatus !== "completed"
      : itemStatus !== "completed" && itemStatus !== "incomplete")
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  if (node.type === "message") {
    return validateResponsesMessage(node, itemStatus);
  }
  if (node.type === "function_call") {
    return validateResponsesFunctionCall(node, itemStatus, tooling);
  }
  if (node.type === "custom_tool_call") {
    return validateResponsesCustomCall(node, itemStatus, tooling);
  }
  return { kind: "rejected", reason: "invalid_envelope" };
}

function validateResponsesMessage(
  node: JsonObject,
  status: SimulatedOutputItemStatus,
): { kind: "accepted"; item: SimulatedOutputItem } | SimulatedOutputRejected {
  if (
    node.role !== "assistant" ||
    !Array.isArray(node.content)
  ) {
    return { kind: "rejected", reason: "invalid_message" };
  }

  const parts: SimulatedOutputMessagePart[] = [];
  let text = "";
  let refusal: string | null = null;
  for (const part of node.content) {
    if (!isJsonObject(part)) {
      return { kind: "rejected", reason: "invalid_message" };
    }
    if (part.type === "output_text" && typeof part.text === "string") {
      parts.push({ kind: "output_text", text: part.text });
      text += part.text;
      continue;
    }
    if (part.type === "refusal" && typeof part.refusal === "string") {
      parts.push({ kind: "refusal", text: part.refusal });
      refusal = refusal === null ? part.refusal : refusal + part.refusal;
      continue;
    }
    return { kind: "rejected", reason: "invalid_message" };
  }
  if (!text.trim() && !refusal?.trim()) {
    return { kind: "rejected", reason: "invalid_message" };
  }
  return {
    kind: "accepted",
    item: { kind: "message", status, content: text, refusal, parts },
  };
}

function validateResponsesFunctionCall(
  node: JsonObject,
  status: SimulatedOutputItemStatus,
  tooling: OpenAiTooling,
): { kind: "accepted"; item: SimulatedOutputItem } | SimulatedOutputRejected {
  if (
    !isNonEmptyString(node.call_id) ||
    !isNonEmptyString(node.name) ||
    typeof node.arguments !== "string"
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const result = validateFunctionCall(
    node.call_id,
    node.name,
    node.arguments,
    tooling,
  );
  if (result.kind === "rejected") {
    return result;
  }
  return {
    kind: "accepted",
    item: { kind: "function_call", status, call: result.call },
  };
}

function validateResponsesCustomCall(
  node: JsonObject,
  status: SimulatedOutputItemStatus,
  tooling: OpenAiTooling,
): { kind: "accepted"; item: SimulatedOutputItem } | SimulatedOutputRejected {
  if (
    !isNonEmptyString(node.call_id) ||
    !isNonEmptyString(node.name) ||
    typeof node.input !== "string"
  ) {
    return { kind: "rejected", reason: "invalid_envelope" };
  }
  const result = validateCustomCall(node.call_id, node.name, node.input, tooling);
  if (result.kind === "rejected") {
    return result;
  }
  return {
    kind: "accepted",
    item: { kind: "custom_tool_call", status, call: result.call },
  };
}

function isResponseStatus(
  value: JsonValue | undefined,
): value is SimulatedResponseStatus {
  return (
    value === "completed" ||
    value === "failed" ||
    value === "incomplete"
  );
}

function isItemStatus(
  value: JsonValue | undefined,
): value is SimulatedOutputItemStatus {
  return (
    value === "in_progress" ||
    value === "completed" ||
    value === "incomplete"
  );
}

function validateFunctionCall(
  id: string,
  rawName: string,
  argumentsJson: string,
  tooling: OpenAiTooling,
): { kind: "accepted"; call: OpenAiAssistantToolCall } | SimulatedOutputRejected {
  if (tooling.tools.some((tool) => tool.name === rawName && tool.type !== "function")) {
    return { kind: "rejected", reason: "tool_type_mismatch" };
  }
  const offered = resolveOfferedTool(rawName, tooling, "function");
  if (!offered) {
    return {
      kind: "rejected",
      reason: rawName.includes(".")
        ? "tool_namespace_mismatch"
        : "unknown_tool",
    };
  }
  if (offered.type !== "function") {
    return { kind: "rejected", reason: "tool_type_mismatch" };
  }
  const call = {
    id,
    name: offered.name,
    type: "function" as const,
    argumentsJson: escapeJsonControlCharactersInStrings(argumentsJson),
  };
  const validation = validateOpenAiToolCall(call, tooling);
  return validation.valid
    ? { kind: "accepted", call }
    : mapToolValidation(validation);
}

function validateCustomCall(
  id: string,
  rawName: string,
  input: string,
  tooling: OpenAiTooling,
): { kind: "accepted"; call: OpenAiAssistantToolCall } | SimulatedOutputRejected {
  if (tooling.tools.some((tool) => tool.name === rawName && tool.type !== "custom")) {
    return { kind: "rejected", reason: "tool_type_mismatch" };
  }
  const offered = resolveOfferedTool(rawName, tooling, "custom");
  if (!offered) {
    return {
      kind: "rejected",
      reason: rawName.includes(".")
        ? "tool_namespace_mismatch"
        : "unknown_tool",
    };
  }
  if (offered.type !== "custom") {
    return { kind: "rejected", reason: "tool_type_mismatch" };
  }
  const call = {
    id,
    name: offered.name,
    type: "custom" as const,
    argumentsJson: input,
  };
  const validation = validateOpenAiToolCall(call, tooling);
  return validation.valid
    ? { kind: "accepted", call }
    : mapToolValidation(validation);
}

function resolveOfferedTool(
  rawName: string,
  tooling: OpenAiTooling,
  type: "function" | "custom",
): OpenAiTooling["tools"][number] | null {
  const exact = tooling.tools.find(
    (tool) => tool.type === type && tool.name === rawName,
  );
  if (exact) {
    return exact;
  }
  const separator = rawName.indexOf(".");
  if (separator <= 0 || separator === rawName.length - 1) {
    return null;
  }
  const namespace = rawName.slice(0, separator);
  const unqualified = rawName.slice(separator + 1);
  return tooling.tools.find(
    (tool) =>
      tool.type === type &&
      tool.name === unqualified &&
      (tool.namespace === namespace ||
        (type === "function" &&
          namespace === functionNamespacePrefix.slice(0, -1) &&
          !tool.namespace)),
  ) ?? null;
}

function mapToolValidation(
  validation: Exclude<OpenAiToolCallValidation, { valid: true }>,
): SimulatedOutputRejected {
  if (validation.reason === "unoffered_tool") {
    return { kind: "rejected", reason: "unknown_tool" };
  }
  if (validation.reason === "tool_choice_mismatch") {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  return { kind: "rejected", reason: validation.reason };
}

function validateToolPolicy(
  calls: SimulatedOutputItem[],
  tooling: OpenAiTooling,
  refusal: string | null,
  status: SimulatedResponseStatus,
): SimulatedOutputRejected | null {
  const callItems = calls.filter(isCallItem);
  if (!isSupportedToolChoice(tooling.toolChoiceMode)) {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  if (!tooling.parallelToolCalls && callItems.length > 1) {
    return { kind: "rejected", reason: "parallel_calls_violation" };
  }
  if (
    (tooling.tools.length === 0 ||
      tooling.toolChoiceMode === ToolChoiceModes.None) &&
    callItems.length > 0
  ) {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  if (status !== "completed") {
    return null;
  }
  const requiresCall =
    tooling.requiredByLocalAction === true ||
    tooling.toolChoiceMode === ToolChoiceModes.Required ||
    tooling.toolChoiceMode === ToolChoiceModes.Function;
  if (requiresCall && callItems.length === 0) {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  if (
    tooling.toolChoiceMode === ToolChoiceModes.Function &&
    tooling.toolChoiceFunctionName &&
    callItems.some(
      (item) =>
        item.call.name !== tooling.toolChoiceFunctionName ||
        (tooling.toolChoiceToolType &&
          item.call.type !== tooling.toolChoiceToolType),
    )
  ) {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  if (requiresCall && refusal?.trim()) {
    return { kind: "rejected", reason: "tool_choice_violation" };
  }
  return null;
}

function isSupportedToolChoice(value: string): boolean {
  return (
    value === ToolChoiceModes.Auto ||
    value === ToolChoiceModes.None ||
    value === ToolChoiceModes.Required ||
    value === ToolChoiceModes.Function
  );
}

function readChatMessage(
  message: JsonObject,
):
  | { kind: "accepted"; item: SimulatedOutputItem | null; refusal: string | null }
  | SimulatedOutputRejected {
  const refusalValue = message.refusal;
  if (
    refusalValue !== undefined &&
    refusalValue !== null &&
    typeof refusalValue !== "string"
  ) {
    return { kind: "rejected", reason: "invalid_message" };
  }

  const content = message.content;
  const parts: SimulatedOutputMessagePart[] = [];
  let text = "";
  if (typeof content === "string") {
    if (content) {
      parts.push({ kind: "output_text", text: content });
      text = content;
    }
  } else if (content !== null && content !== undefined) {
    if (!Array.isArray(content)) {
      return { kind: "rejected", reason: "invalid_message" };
    }
    for (const part of content) {
      if (!isJsonObject(part)) {
        return { kind: "rejected", reason: "invalid_message" };
      }
      if (part.type === "text" && typeof part.text === "string") {
        parts.push({ kind: "output_text", text: part.text });
        text += part.text;
        continue;
      }
      if (part.type === "refusal" && typeof part.refusal === "string") {
        parts.push({ kind: "refusal", text: part.refusal });
        continue;
      }
      return { kind: "rejected", reason: "invalid_message" };
    }
  }

  const refusal = typeof refusalValue === "string"
    ? refusalValue
    : parts
        .filter((part): part is Extract<SimulatedOutputMessagePart, { kind: "refusal" }> =>
          part.kind === "refusal"
        )
        .map((part) => part.text)
        .join("");
  if (refusal && content !== null && content !== undefined && parts.every(
    (part) => part.kind !== "output_text",
  )) {
    return {
      kind: "accepted",
      item: { kind: "message", status: "completed", content: "", refusal, parts },
      refusal,
    };
  }
  if (!text.trim() && !refusal.trim()) {
    return {
      kind: "accepted",
      item: null,
      refusal: refusal || null,
    };
  }
  return {
    kind: "accepted",
    item: {
      kind: "message",
      status: "completed",
      content: text,
      refusal: refusal || null,
      parts,
    },
    refusal: refusal || null,
  };
}

function responseText(items: SimulatedOutputItem[]): string {
  return items
    .filter((item): item is Extract<SimulatedOutputItem, { kind: "message" }> =>
      item.kind === "message"
    )
    .map((item) => item.content)
    .join("");
}

function isCallItem(
  item: SimulatedOutputItem,
): item is Exclude<SimulatedOutputItem, { kind: "message" }> {
  return item.kind === "function_call" || item.kind === "custom_tool_call";
}

function isNonEmptyString(value: JsonValue | undefined): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

export function buildLocalChatCompletion(
  accepted: SimulatedOutputAccepted,
  model: string,
  conversationId: string,
  includeConversationId: boolean,
  rawRequest?: JsonObject,
): JsonObject {
  const messageItem = accepted.items.find((item) => item.kind === "message");
  const calls = accepted.items.filter(isCallItem);
  const message: JsonObject = {
    role: "assistant",
    content: messageItem
      ? messageItem.refusal
        ? null
        : messageItem.content
      : null,
  };
  if (messageItem?.refusal) {
    message.refusal = messageItem.refusal;
  }
  if (calls.length > 0) {
    message.tool_calls = calls.map((item) => ({
      id: item.call.id,
      type: "function",
      function: {
        name: item.call.name,
        arguments: item.call.argumentsJson,
      },
    }));
  }

  const response: JsonObject = {
    id: `chatcmpl-${cryptoRandomUuid()}`,
    object: "chat.completion",
    created: nowUnix(),
    model,
    choices: [{
      index: 0,
      message,
      finish_reason: accepted.finishReason ?? (calls.length > 0 ? "tool_calls" : "stop"),
    }],
  };
  if (includeConversationId) {
    response.conversation_id = conversationId;
  }
  response.usage = buildLocalChatUsage(rawRequest, response);
  return response;
}

export function buildLocalChatCompletionChunk(
  completionId: string,
  created: number,
  model: string,
  conversationId: string | null,
  role: string | null,
  content: string | null,
  refusal: string | null,
  finishReason: string | null,
  calls?: OpenAiAssistantToolCall[],
  usage?: JsonObject,
): JsonObject {
  const delta: JsonObject = {};
  if (role) delta.role = role;
  if (content !== null) delta.content = content;
  if (refusal !== null) delta.refusal = refusal;
  if (calls && calls.length > 0) {
    delta.tool_calls = calls.map((call, index) => ({
      index,
      id: call.id,
      type: "function",
      function: {
        name: call.name,
        arguments: call.argumentsJson,
      },
    }));
  }
  const chunk: JsonObject = {
    id: completionId,
    object: "chat.completion.chunk",
    created,
    model,
    choices: usage
      ? []
      : [{ index: 0, delta, finish_reason: finishReason }],
  };
  if (usage) {
    chunk.usage = usage;
  }
  if (conversationId) chunk.conversation_id = conversationId;
  return chunk;
}

export function buildLocalResponsesProjection(
  accepted: SimulatedOutputAccepted,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  includeConversationId: boolean,
  responseId = createOpenAiResponseId(),
  createdAt = nowUnix(),
): {
  responseId: string;
  createdAt: number;
  outputItems: JsonObject[];
  responseBody: JsonObject;
} {
  const outputItems = accepted.items.map((item) =>
    buildLocalResponsesOutputItem(item),
  );
  const terminal = buildLocalTerminal(
    accepted.status,
    accepted.terminalCause,
  );
  const responseBody = buildOpenAiResponseObject(
    responseId,
    createdAt,
    parsedRequest.base.model,
    accepted.status,
    outputItems,
    parsedRequest,
    includeConversationId ? conversationId : null,
    undefined,
    terminal,
  );
  return { responseId, createdAt, outputItems, responseBody };
}

function buildLocalResponsesOutputItem(
  item: SimulatedOutputItem,
): JsonObject {
  if (item.kind === "message") {
    const content: JsonValue[] = [];
    for (const part of item.parts) {
      if (part.kind === "output_text") {
        content.push({
          type: "output_text",
          text: stripPrivateCitationMarkers(part.text),
          annotations: [],
        });
      } else {
        content.push({ type: "refusal", refusal: part.text });
      }
    }
    return {
      id: createOpenAiOutputItemId("msg"),
      type: "message",
      status: item.status,
      role: "assistant",
      content,
    };
  }
  return buildFunctionCallOutputItem(
    createOpenAiOutputItemId(item.kind === "custom_tool_call" ? "ctc" : "fc"),
    item.call,
    item.status,
  );
}

function buildLocalChatUsage(
  rawRequest: JsonObject | undefined,
  response: JsonObject,
): JsonObject {
  const requestEnvelope = rawRequest ? { ...rawRequest } : {};
  delete requestEnvelope.stream;
  delete requestEnvelope.stream_options;
  const promptTokens = estimateJsonTokens(JSON.stringify(requestEnvelope));
  const completionTokens = estimateJsonTokens(
    JSON.stringify(response.choices ?? []),
  );
  return {
    prompt_tokens: promptTokens,
    completion_tokens: completionTokens,
    total_tokens: promptTokens + completionTokens,
  };
}

function buildLocalTerminal(
  status: SimulatedResponseStatus,
  cause: SimulatedOutputTerminalCause | null,
): ResponseTerminal | undefined {
  if (status === "completed") {
    return undefined;
  }
  if (status === "failed") {
    return {
      status,
      error: {
        code: cause?.kind === "failed"
          ? cause.code
          : "simulated_provider_failure",
        message: cause?.kind === "failed"
          ? cause.message
          : "The simulated provider returned a failed response.",
      },
      incomplete_details: null,
    };
  }
  return {
    status,
    error: null,
    incomplete_details: {
      reason: cause?.kind === "incomplete"
        ? cause.reason
        : "simulated_provider_incomplete",
    },
  };
}

function cryptoRandomUuid(): string {
  return randomUUID().replaceAll("-", "");
}
