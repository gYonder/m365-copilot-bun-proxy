import { randomUUID } from "node:crypto";
import type {
  JsonObject,
  JsonValue,
  OpenAiAssistantResponse,
  OpenAiAssistantToolCall,
  ParsedResponsesRequest,
} from "./types";
import { cloneJsonValue } from "./utils";

export function createOpenAiResponseId(): string {
  return `resp_${randomUUID().replaceAll("-", "")}`;
}

export function createOpenAiOutputItemId(prefix: string): string {
  return `${prefix}_${randomUUID().replaceAll("-", "")}`;
}

export function buildOpenAiResponseFromAssistant(
  responseId: string,
  createdAt: number,
  model: string,
  status: string,
  parsedRequest: ParsedResponsesRequest,
  assistantResponse: OpenAiAssistantResponse,
  includeConversationId: boolean,
  conversationId: string,
): JsonObject {
  const output =
    assistantResponse.toolCalls.length > 0
      ? buildFunctionCallOutputItems(assistantResponse.toolCalls, "completed")
      : [
          buildMessageOutputItem(
            createOpenAiOutputItemId("msg"),
            assistantResponse.content ?? "",
            "completed",
          ),
        ];

  return buildOpenAiResponseObject(
    responseId,
    createdAt,
    model,
    status,
    output,
    parsedRequest,
    includeConversationId ? conversationId : null,
  );
}

export function buildOpenAiResponseObject(
  responseId: string,
  createdAt: number,
  model: string,
  status: string,
  output: JsonObject[],
  parsedRequest: ParsedResponsesRequest,
  conversationId: string | null,
): JsonObject {
  const usage = buildResponseUsage(parsedRequest, output);
  const response: JsonObject = {
    id: responseId,
    object: "response",
    created_at: createdAt,
    status,
    error: null,
    incomplete_details: null,
    service_tier: null,
    usage,
    model,
    output: output.map((item) => cloneJsonValue(item)),
    output_text: extractOutputText(output),
    parallel_tool_calls: parsedRequest.base.tooling.parallelToolCalls,
    store: parsedRequest.store,
  };

  response.previous_response_id = parsedRequest.previousResponseId ?? null;
  response.instructions = parsedRequest.instructions ?? null;
  if (parsedRequest.inputItemsForStorage.length > 0) {
    response.input = cloneJsonValue(parsedRequest.inputItemsForStorage);
  }
  if (conversationId?.trim()) {
    response.conversation = conversationId;
    response.conversation_id = conversationId;
  }

  return response;
}

export function buildMessageOutputItem(
  itemId: string,
  text: string,
  status: string,
): JsonObject {
  return {
    id: itemId,
    type: "message",
    status,
    role: "assistant",
    content: [{ type: "output_text", text, annotations: [] }],
  };
}

export function buildFunctionCallOutputItems(
  toolCalls: OpenAiAssistantToolCall[],
  status: string,
): JsonObject[] {
  return toolCalls.map((toolCall) =>
    buildFunctionCallOutputItem(
      createOpenAiOutputItemId("fc"),
      toolCall,
      status,
    ),
  );
}

export function buildFunctionCallOutputItem(
  itemId: string,
  toolCall: OpenAiAssistantToolCall,
  status: string,
): JsonObject {
  return {
    id: itemId,
    type: toolCall.type === "custom" ? "custom_tool_call" : "function_call",
    status,
    call_id: toolCall.id,
    name: toolCall.name,
    ...(toolCall.type === "custom"
      ? { input: toolCall.argumentsJson }
      : { arguments: toolCall.argumentsJson }),
  };
}

export function buildResponseCreatedEvent(response: JsonObject): JsonObject {
  return { type: "response.created", response: cloneJsonValue(response) };
}

export function buildResponseInProgressEvent(response: JsonObject): JsonObject {
  return { type: "response.in_progress", response: cloneJsonValue(response) };
}

export function buildResponseOutputItemAddedEvent(
  responseId: string,
  outputIndex: number,
  item: JsonObject,
): JsonObject {
  return {
    type: "response.output_item.added",
    response_id: responseId,
    output_index: outputIndex,
    item: cloneJsonValue(item),
  };
}

export function buildResponseContentPartAddedEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  part: JsonObject,
  contentIndex = 0,
): JsonObject {
  return {
    type: "response.content_part.added",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    content_index: contentIndex,
    part: cloneJsonValue(part),
  };
}

export function buildResponseOutputTextDeltaEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  delta: string,
  contentIndex = 0,
): JsonObject {
  return {
    type: "response.output_text.delta",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    content_index: contentIndex,
    delta,
  };
}

export function buildResponseOutputTextDoneEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  text: string,
  contentIndex = 0,
): JsonObject {
  return {
    type: "response.output_text.done",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    content_index: contentIndex,
    text,
  };
}

export function buildResponseContentPartDoneEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  part: JsonObject,
  contentIndex = 0,
): JsonObject {
  return {
    type: "response.content_part.done",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    content_index: contentIndex,
    part: cloneJsonValue(part),
  };
}

export function buildResponseOutputItemDoneEvent(
  responseId: string,
  outputIndex: number,
  item: JsonObject,
): JsonObject {
  return {
    type: "response.output_item.done",
    response_id: responseId,
    output_index: outputIndex,
    item: cloneJsonValue(item),
  };
}

export function buildResponseCompletedEvent(response: JsonObject): JsonObject {
  return { type: "response.completed", response: cloneJsonValue(response) };
}

function extractOutputText(output: JsonObject[]): string {
  const segments: string[] = [];
  for (const item of output) {
    if ((item.type ?? "") !== "message") {
      continue;
    }
    const content = item.content;
    if (!Array.isArray(content)) {
      continue;
    }
    for (const contentItem of content) {
      if (
        !contentItem ||
        typeof contentItem !== "object" ||
        Array.isArray(contentItem)
      ) {
        continue;
      }
      const typed = contentItem as Record<string, unknown>;
      if ((typed.type ?? "") !== "output_text") {
        continue;
      }
      const text = typed.text;
      if (typeof text === "string" && text.length > 0) {
        segments.push(text);
      }
    }
  }
  return segments.join("");
}

function buildResponseUsage(
  parsedRequest: ParsedResponsesRequest,
  output: JsonObject[],
): JsonObject {
  const inputText = extractInputText(parsedRequest.inputItemsForStorage);
  const outputText = extractOutputText(output);
  const inputTokens = estimateTokenCount(inputText);
  const outputTokens = estimateTokenCount(outputText);
  return {
    input_tokens: inputTokens,
    input_tokens_details: { cached_tokens: 0 },
    output_tokens: outputTokens,
    output_tokens_details: { reasoning_tokens: 0 },
    total_tokens: inputTokens + outputTokens,
  };
}

function extractInputText(inputItems: JsonValue[]): string {
  const segments: string[] = [];
  for (const inputItem of inputItems) {
    if (
      !inputItem ||
      typeof inputItem !== "object" ||
      Array.isArray(inputItem)
    ) {
      continue;
    }
    const record = inputItem as Record<string, unknown>;
    const content = record.content;
    if (typeof content === "string") {
      if (content.trim().length > 0) {
        segments.push(content);
      }
      continue;
    }
    if (!Array.isArray(content)) {
      continue;
    }
    for (const part of content) {
      if (
        !part ||
        typeof part !== "object" ||
        Array.isArray(part)
      ) {
        continue;
      }
      const partRecord = part as Record<string, unknown>;
      const text = partRecord.text;
      if (typeof text === "string" && text.trim().length > 0) {
        segments.push(text);
      }
    }
  }
  return segments.join("\n");
}

function estimateTokenCount(text: string): number {
  const normalized = text.trim();
  if (!normalized) {
    return 0;
  }
  // Approximate token count for compatibility fields; precise tokenizer not required here.
  return Math.max(1, Math.ceil(normalized.length / 4));
}
