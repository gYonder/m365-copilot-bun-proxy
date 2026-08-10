import {
  OpenAiTransformModes,
  ResponseFormatTypes,
  ToolChoiceModes,
  TransportNames,
  type ContextMessage,
  type JsonObject,
  type JsonValue,
  type OpenAiResponseFormat,
  type OpenAiToolDefinition,
  type OpenAiTooling,
  type ParsedOpenAiRequest,
  type ParsedImageInput,
  type ParsedResponsesRequest,
  type ResponsesProtocolIdentity,
  type WrapperOptions,
} from "./types";
import {
  firstNonEmpty,
  normalizeNullableString,
  parseBooleanString,
  tryGetBoolean,
  tryGetDouble,
  tryGetString,
  isJsonObject,
  cloneJsonValue,
  tryParseJsonObject,
} from "./utils";
import { parseImageInputs } from "./image-input";

const MAX_SIMULATED_TOOL_RESULT_CHARS = 40_000;

export function normalizeTransport(
  transport: string | null | undefined,
): string {
  if (!transport || !transport.trim()) {
    return TransportNames.Graph;
  }
  return transport.trim().toLowerCase();
}

export function normalizeOpenAiTransformMode(
  mode: string | null | undefined,
): string {
  if (!mode || !mode.trim()) {
    return OpenAiTransformModes.Simulated;
  }
  const normalized = mode.trim().toLowerCase();
  return normalized === OpenAiTransformModes.Mapped
    ? OpenAiTransformModes.Mapped
    : OpenAiTransformModes.Simulated;
}

export function isSimulatedOpenAiTransformMode(mode: string): boolean {
  return normalizeOpenAiTransformMode(mode) === OpenAiTransformModes.Simulated;
}

export function isSupportedOpenAiTransformMode(
  mode: string | null | undefined,
): boolean {
  if (!mode || !mode.trim()) {
    return false;
  }
  const normalized = mode.trim().toLowerCase();
  return (
    normalized === OpenAiTransformModes.Simulated ||
    normalized === OpenAiTransformModes.Mapped
  );
}

export function isSupportedTransport(
  transport: string | null | undefined,
): boolean {
  const normalized = normalizeTransport(transport);
  return (
    normalized === TransportNames.Graph ||
    normalized === TransportNames.Substrate
  );
}

export function resolveTransport(
  request: Request,
  requestJson: JsonObject,
  options: WrapperOptions,
): string {
  const transport = firstNonEmpty(
    request.headers.get("x-m365-transport"),
    tryGetString(requestJson, "m365_transport"),
    tryGetString(requestJson, "transport"),
    options.transport,
  );
  return normalizeTransport(transport);
}

export function selectConversation(
  request: Request,
  requestJson: JsonObject,
  fallbackConversationKey: string | null,
): {
  conversationId: string | null;
  conversationKey: string | null;
  forceNewConversation: boolean;
} {
  const specConversationId = extractSpecConversationId(requestJson);
  const conversationId = normalizeNullableString(
    firstNonEmpty(
      request.headers.get("x-m365-conversation-id"),
      specConversationId,
      tryGetString(requestJson, "m365_conversation_id"),
      tryGetString(requestJson, "conversation_id"),
    ),
  );

  const conversationKey = normalizeNullableString(
    firstNonEmpty(
      request.headers.get("x-m365-conversation-key"),
      tryGetString(requestJson, "m365_conversation_key"),
      fallbackConversationKey,
    ),
  );

  const forceNewFromHeader = parseBooleanString(
    request.headers.get("x-m365-new-conversation"),
  );
  const forceNewFromBody =
    tryGetBoolean(requestJson, "m365_new_conversation") === true;

  return {
    conversationId,
    conversationKey,
    forceNewConversation: forceNewFromHeader || forceNewFromBody,
  };
}

function extractSpecConversationId(requestJson: JsonObject): string | null {
  const conversation = requestJson.conversation;
  if (typeof conversation === "string") {
    return conversation;
  }
  if (!isJsonObject(conversation)) {
    return null;
  }
  return tryGetString(conversation, "id");
}

export function scopeConversationKey(
  conversationKey: string | null,
  transportName: string,
  compatibilityKey: string | null = null,
): string | null {
  const normalizedKey = normalizeNullableString(conversationKey);
  if (!normalizedKey) {
    return null;
  }
  const normalizedCompatibilityKey = normalizeNullableString(compatibilityKey);
  return normalizedCompatibilityKey
    ? `${normalizeTransport(transportName)}:${normalizedKey}:profile:${normalizedCompatibilityKey}`
    : `${normalizeTransport(transportName)}:${normalizedKey}`;
}

export function buildCopilotRequestPayload(
  request: ParsedOpenAiRequest,
): JsonObject {
  const payload: JsonObject = {
    message: { text: request.promptText },
    locationHint: cloneJsonValue(request.locationHint),
  };

  if (request.additionalContext.length > 0) {
    const additionalContext: JsonValue[] = [];
    for (const item of request.additionalContext) {
      if (!item.text.trim()) {
        continue;
      }
      const context: JsonObject = { text: item.text };
      if (item.description?.trim()) {
        context.description = item.description;
      }
      additionalContext.push(context);
    }
    if (additionalContext.length > 0) {
      payload.additionalContext = additionalContext;
    }
  }

  if (request.contextualResources) {
    payload.contextualResources = cloneJsonValue(request.contextualResources);
  }
  return payload;
}

export function tryParseOpenAiRequest(
  requestJson: JsonObject,
  options: WrapperOptions,
): { ok: true; request: ParsedOpenAiRequest } | { ok: false; error: string } {
  const transformMode = normalizeOpenAiTransformMode(options.openAiTransformMode);
  if (transformMode === OpenAiTransformModes.Simulated) {
    const imageResult = parseCurrentTurnImages(requestJson, options);
    if (!imageResult.ok) {
      return { ok: false, error: imageResult.message };
    }
    return {
      ok: true,
      request: buildSimulatedOpenAiRequest(
        requestJson,
        options,
        "chat.completions",
        imageResult.images,
      ),
    };
  }

  const messagesNode = requestJson.messages;
  if (!Array.isArray(messagesNode) || messagesNode.length === 0) {
    return {
      ok: false,
      error: "The 'messages' array is required and cannot be empty.",
    };
  }

  const tooling = requireLocalToolingIfNeeded(
    requestJson,
    parseTooling(requestJson),
  );
  const responseFormat = parseResponseFormat(requestJson);
  const reasoningEffort = tryGetString(requestJson, "reasoning_effort");
  const temperature = tryGetDouble(requestJson, "temperature");

  const messages: {
    role: string;
    content: string;
    index: number;
    hasImage: boolean;
  }[] = [];
  for (let index = 0; index < messagesNode.length; index++) {
    const message = messagesNode[index];
    if (!isJsonObject(message)) {
      continue;
    }

    const role = (tryGetString(message, "role") ?? "user").toLowerCase();
    let content = extractMessageContent(message.content);
    const hasImage = hasImageContent(message.content);

    if (role === "assistant") {
      const assistantToolCalls = tryExtractAssistantToolCalls(message, content);
      if (assistantToolCalls.length > 0) {
        content = convertToolCallsToContextText(assistantToolCalls);
      }
    }

    if (!content && role === "tool") {
      const toolName = tryGetString(message, "name");
      const toolCallId = tryGetString(message, "tool_call_id");
      let prefix = toolName ? `tool:${toolName}` : "tool";
      if (toolCallId) {
        prefix += `[${toolCallId}]`;
      }
      const toolPayload = stringifyJsonValue(message.content);
      if (toolPayload) {
        content = `${prefix}: ${toolPayload}`;
      }
    }

    if (!content.trim() && !hasImage) {
      continue;
    }
    messages.push({ role, content: content.trim(), index, hasImage });
  }

  if (messages.length === 0) {
    return {
      ok: false,
      error: "No textual content could be extracted from 'messages'.",
    };
  }

  const prompt = resolvePrompt(messages);
  if (!prompt || (!prompt.content.trim() && !prompt.hasImage)) {
    return {
      ok: false,
      error: "Unable to determine a prompt from the message list.",
    };
  }
  const promptMessage = messagesNode[prompt.index];
  const imageResult = parseImageInputs(
    isJsonObject(promptMessage) ? promptMessage.content : undefined,
    imageInputLimits(options),
  );
  if (!imageResult.ok) {
    return { ok: false, error: imageResult.message };
  }

  const stream = tryGetBoolean(requestJson, "stream") === true;
  const model =
    tryGetString(requestJson, "model") ||
    (options.defaultModel?.trim() ? options.defaultModel : "m365-copilot");

  const parsedRequest: ParsedOpenAiRequest = {
    model,
    stream,
    transformMode,
    hostedWebSearch:
      !tooling.requiredByLocalAction && hasNativeWebSearchIntent(requestJson),
    promptText: buildMappedPromptText(prompt.content, tooling),
    userKey: tryGetString(requestJson, "user"),
    locationHint: buildLocationHint(requestJson, options.defaultTimeZone),
    contextualResources: buildContextualResources(requestJson),
    additionalContext: buildAdditionalContext(
      messages,
      prompt.index,
      requestJson,
      options.maxAdditionalContextMessages,
      tooling,
      responseFormat,
      reasoningEffort,
      temperature,
    ),
    tooling,
    responseFormat,
    reasoningEffort,
    temperature,
    images: imageResult.images,
  };

  return { ok: true, request: parsedRequest };
}

export function tryParseResponsesRequest(
  requestJson: JsonObject,
  options: WrapperOptions,
):
  | { ok: true; request: ParsedResponsesRequest }
  | { ok: false; error: string } {
  const previousResponseId = tryGetString(requestJson, "previous_response_id");
  const protocolIdentity = extractResponsesProtocolIdentity(
    requestJson,
    previousResponseId,
  );
  const contextWindowId = extractContextWindowId(requestJson);
  const store = tryGetBoolean(requestJson, "store") !== false;
  const specConversationId = extractSpecConversationId(requestJson);
  if (previousResponseId && specConversationId) {
    return {
      ok: false,
      error:
        "The 'conversation' field cannot be used together with 'previous_response_id'.",
    };
  }

  const transformMode = normalizeOpenAiTransformMode(options.openAiTransformMode);
  if (transformMode === OpenAiTransformModes.Simulated) {
    const normalizedInput = normalizeResponsesInput(requestJson.input);
    const imageResult = parseCurrentTurnImages(requestJson, options);
    if (!imageResult.ok) {
      return { ok: false, error: imageResult.message };
    }
    return {
      ok: true,
      request: {
        base: buildSimulatedOpenAiRequest(
          requestJson,
          options,
          "responses",
          imageResult.images,
        ),
        previousResponseId,
        protocolIdentity,
        inputItemsForStorage: normalizedInput.inputItemsForStorage,
        instructions: tryGetString(requestJson, "instructions"),
        store,
        rawRequest: cloneJsonValue(requestJson),
        contextWindowId,
        contextInputTokens: null,
      },
    };
  }

  const normalized = cloneJsonValue(requestJson);
  const normalizedInput = normalizeResponsesInput(requestJson.input);
  if (normalizedInput.messages.length === 0) {
    return {
      ok: false,
      error:
        "The 'input' field is required and must contain at least one supported message item.",
    };
  }

  normalized.messages = normalizedInput.messages;

  if (normalized.response_format === undefined) {
    const mappedResponseFormat = mapResponsesTextFormat(requestJson);
    if (mappedResponseFormat) {
      normalized.response_format = mappedResponseFormat;
    }
  }

  if (normalized.reasoning_effort === undefined) {
    const reasoningEffort = mapResponsesReasoningEffort(requestJson);
    if (reasoningEffort) {
      normalized.reasoning_effort = reasoningEffort;
    }
  }

  const instructions = tryGetString(requestJson, "instructions");
  if (instructions && !tryGetString(normalized, "m365_system_prompt")) {
    normalized.m365_system_prompt = instructions;
  }

  const parsedBase = tryParseOpenAiRequest(normalized, options);
  if (!parsedBase.ok) {
    return parsedBase;
  }

  return {
    ok: true,
    request: {
      base: parsedBase.request,
      previousResponseId,
      protocolIdentity,
      inputItemsForStorage: normalizedInput.inputItemsForStorage,
      instructions,
      store,
      rawRequest: cloneJsonValue(requestJson),
      contextWindowId,
      contextInputTokens: null,
    },
  };
}

function extractResponsesProtocolIdentity(
  requestJson: JsonObject,
  previousResponseId: string | null,
): ResponsesProtocolIdentity {
  const metadata = isJsonObject(requestJson.client_metadata)
    ? requestJson.client_metadata
    : null;
  const encoded = tryGetString(metadata, "x-codex-turn-metadata");
  const encodedMetadata = encoded ? tryParseJsonObject(encoded) : null;
  const readMetadata = (key: string): string | null =>
    tryGetString(metadata, key) ?? tryGetString(encodedMetadata, key);
  const callIds = new Set<string>();
  const input = requestJson.input;
  if (Array.isArray(input)) {
    for (const item of input) {
      if (!isJsonObject(item)) continue;
      const type = (tryGetString(item, "type") ?? "").toLowerCase();
      if (
        type !== "function_call_output" &&
        type !== "custom_tool_call_output"
      ) continue;
      const callId = tryGetString(item, "call_id");
      if (callId) callIds.add(callId);
    }
  }
  return {
    threadId: readMetadata("thread_id"),
    sessionId: readMetadata("session_id"),
    turnId: readMetadata("turn_id"),
    conversationId: extractSpecConversationId(requestJson),
    previousResponseId,
    callIds: [...callIds].sort(),
  };
}

function extractContextWindowId(requestJson: JsonObject): string | null {
  const metadata = isJsonObject(requestJson.client_metadata)
    ? requestJson.client_metadata
    : null;
  const direct = tryGetString(metadata, "x-codex-window-id");
  if (direct) return direct;
  const encoded = tryGetString(metadata, "x-codex-turn-metadata");
  if (!encoded) return null;
  const parsed = tryParseJsonObject(encoded);
  return tryGetString(parsed, "window_id");
}

function buildSimulatedOpenAiRequest(
  requestJson: JsonObject,
  options: WrapperOptions,
  endpointFormat: "chat.completions" | "responses",
  images: ParsedImageInput[] = [],
): ParsedOpenAiRequest {
  const tooling = requireLocalToolingIfNeeded(
    requestJson,
    parseTooling(requestJson),
  );
  const model =
    tryGetString(requestJson, "model") ||
    (options.defaultModel?.trim() ? options.defaultModel : "m365-copilot");

  return {
    model,
    stream: tryGetBoolean(requestJson, "stream") === true,
    transformMode: OpenAiTransformModes.Simulated,
    hostedWebSearch:
      !tooling.requiredByLocalAction && hasNativeWebSearchIntent(requestJson),
    promptText: buildSimulatedPrompt(
      endpointFormat,
      images.length > 0 ? stripImageInputsFromPrompt(requestJson) : requestJson,
      tooling,
      options.substrate.truncateBeforeSending
        ? options.substrate.maxSendChars
        : 0,
    ),
    userKey: tryGetString(requestJson, "user"),
    locationHint: buildLocationHint(requestJson, options.defaultTimeZone),
    contextualResources: buildContextualResources(requestJson),
    additionalContext: [],
    tooling,
    responseFormat: parseResponseFormat(requestJson),
    reasoningEffort: tryGetString(requestJson, "reasoning_effort"),
    temperature: tryGetDouble(requestJson, "temperature"),
    images,
  };
}

function stripImageInputsFromPrompt(requestJson: JsonObject): JsonObject {
  const sanitized = stripImageValue(requestJson);
  if (isJsonObject(sanitized)) {
    return sanitized;
  }
  return {};
}

function stripImageValue(value: JsonValue): JsonValue | null {
  if (isImageContentPart(value)) {
    return null;
  }
  if (Array.isArray(value)) {
    return value
      .map(stripImageValue)
      .filter((item): item is JsonValue => item !== null);
  }
  if (!isJsonObject(value)) {
    return value;
  }
  const result: JsonObject = {};
  for (const [key, child] of Object.entries(value)) {
    const sanitizedChild = stripImageValue(child);
    if (sanitizedChild !== null) {
      result[key] = sanitizedChild;
    }
  }
  return result;
}

function buildSimulatedPrompt(
  endpointFormat: "chat.completions" | "responses",
  requestJson: JsonObject,
  tooling: OpenAiTooling,
  maxChars = 0,
): string {
  const endpointPath =
    endpointFormat === "responses" ? "/v1/responses" : "/v1/chat/completions";
  const lines: string[] = [
    `The JSON payload below is an entire request for the OpenAI ${endpointFormat} format.`,
    `The JSON payload below is an entire request for POST ${endpointPath}.`,
    `Interpret it exactly in OpenAI ${endpointFormat} format and produce the corresponding response in the same format.`,
    "Focus on producing a valid response object that matches the expected OpenAI format for this request.",
    "Return exactly one markdown JSON code block containing a single valid JSON object and no surrounding prose.",
    'If the payload has "stream": true, still return the final completed JSON object (not SSE events).',
  ];
  const requiresToolCall =
    tooling.tools.length > 0 &&
    (tooling.toolChoiceMode === ToolChoiceModes.Required ||
      tooling.toolChoiceMode === ToolChoiceModes.Function ||
      shouldRequireInitialLocalToolCall(requestJson, tooling));

  if (tooling.tools.length > 0) {
    lines.push(
      "You are producing a response for a local harness that will execute tool calls.",
      "If the request requires local files, shell state, or any other local environment access, emit an appropriate tool call instead of saying the environment is inaccessible.",
      "Do not claim you inspected, changed, or verified local files unless the response includes the matching tool call.",
    );
    if (endpointFormat === "chat.completions") {
      lines.push(
        "Tool calls are supported here: emit assistant tool calls when appropriate.",
        'If returning tool calls, use choices[0].message.tool_calls and set choices[0].finish_reason to "tool_calls".',
        "Do not refuse by saying tool invocation is unsupported.",
        "For each tool call, function.arguments must be a JSON string value (not an object).",
        "For apply_diff calls, each SEARCH block must contain non-empty exact text to match.",
        "If creating/replacing file contents from empty input, prefer write_to_file instead of apply_diff.",
      );
    } else {
      lines.push(
        "Tool calls are supported here: emit function_call output items when appropriate.",
        'If returning tool calls, place them in output items with type "function_call".',
        "Do not refuse by saying tool invocation is unsupported.",
        "For each function_call output item, arguments must be a JSON string value (not an object).",
        "For apply_diff calls, each SEARCH block must contain non-empty exact text to match.",
        "If creating/replacing file contents from empty input, prefer write_to_file instead of apply_diff.",
      );
    }

    if (tooling.toolChoiceMode === ToolChoiceModes.Required) {
      lines.push(
        "This request requires at least one tool call. Do not return a plain-text-only assistant response.",
      );
    } else if (shouldRequireInitialLocalToolCall(requestJson, tooling)) {
      lines.push(
        "The latest user request needs local workspace access and this request has no prior tool result, so this response must include at least one tool call.",
        "Do not return a message-only response for this turn.",
      );
    } else if (
      tooling.toolChoiceMode === ToolChoiceModes.Function &&
      tooling.toolChoiceFunctionName
    ) {
      lines.push(
        `This request requires calling tool "${tooling.toolChoiceFunctionName}".`,
      );
    }
  }

  const trailingContract = requiresToolCall
    ? buildTrailingSimulatedToolContract(endpointFormat, tooling)
    : [];
  const render = (payload: JsonObject, compact: boolean): string =>
    [
      ...lines,
      "```json",
      JSON.stringify(payload, null, compact ? undefined : 2),
      "```",
      ...trailingContract,
    ].join("\n");
  if (maxChars <= 0) {
    return render(requestJson, false);
  }

  function buildTrailingSimulatedToolContract(
    endpointFormat: "chat.completions" | "responses",
    tooling: OpenAiTooling,
  ): string[] {
    const available = tooling.tools
      .map((tool) => `${tool.name} (${tool.type})`)
      .join(", ");
    if (endpointFormat === "responses") {
      return [
        "",
        "CRITICAL OUTPUT CONTRACT — this instruction appears after the request so it has highest priority:",
        "This turn requires a local tool call. Return exactly one markdown JSON code block and no prose.",
        "For a function tool, output one item with this exact shape:",
        '{"id":"resp_1","object":"response","status":"completed","model":"gpt-5.6-sol","output":[{"type":"function_call","call_id":"call_1","name":"TOOL_NAME","arguments":"{\\"key\\":\\"value\\"}"}]}',
        "For a custom tool, use type custom_tool_call and input instead of arguments.",
        `Available local tools: ${available}`,
        "Do not return a message item, output_text, a refusal, or a description of the tool call.",
      ];
    }
    return [
      "",
      "CRITICAL OUTPUT CONTRACT — this instruction appears after the request so it has highest priority:",
      "This turn requires a local tool call. Return exactly one markdown JSON code block and no prose.",
      'Use choices[0].message.tool_calls, JSON-string function.arguments, and finish_reason "tool_calls".',
      `Available local tools: ${available}`,
      "Do not return a plain assistant message, refusal, or description of the tool call.",
    ];
  }
  const compactPrompt = render(requestJson, true);
  const originalCandidates = collectReducibleToolResults(requestJson);
  if (
    compactPrompt.length <= maxChars &&
    originalCandidates.every(
      (candidate) => candidate.value.length <= MAX_SIMULATED_TOOL_RESULT_CHARS,
    )
  ) {
    return compactPrompt;
  }

  // Codex continuations can contain a very large shell/diff result. Keep the
  // request valid JSON and retain the call/result item, but shorten only the
  // result body. The real user message, instructions, tool declarations, call
  // IDs, and the useful head/tail of the result remain intact.
  const reducedRequest = cloneJsonValue(requestJson);
  if (!isJsonObject(reducedRequest)) {
    return compactPrompt;
  }
  const candidates = collectReducibleToolResults(reducedRequest).sort(
    (left, right) => right.value.length - left.value.length,
  );
  for (const candidate of candidates) {
    if (candidate.value.length <= MAX_SIMULATED_TOOL_RESULT_CHARS) continue;
    candidate.owner[candidate.key] = truncateToolResult(
      candidate.value,
      MAX_SIMULATED_TOOL_RESULT_CHARS,
    );
    candidate.value = candidate.owner[candidate.key] as string;
  }
  let reducedPrompt = render(reducedRequest, true);
  for (const candidate of candidates) {
    if (reducedPrompt.length <= maxChars) {
      break;
    }
    const excess = reducedPrompt.length - maxChars;
    const targetChars = Math.max(2_048, candidate.value.length - excess - 128);
    if (targetChars >= candidate.value.length) {
      continue;
    }
    candidate.owner[candidate.key] = truncateToolResult(
      candidate.value,
      targetChars,
    );
    reducedPrompt = render(reducedRequest, true);
  }
  return reducedPrompt;
}

type ReducibleToolResult = {
  owner: JsonObject;
  key: "content" | "output";
  value: string;
};

function collectReducibleToolResults(
  requestJson: JsonObject,
): ReducibleToolResult[] {
  const results: ReducibleToolResult[] = [];
  const input = requestJson.input;
  if (Array.isArray(input)) {
    for (const item of input) {
      if (!isJsonObject(item)) continue;
      const type = (tryGetString(item, "type") ?? "").toLowerCase();
      if (
        (type === "function_call_output" ||
          type === "custom_tool_call_output") &&
        typeof item.output === "string"
      ) {
        results.push({ owner: item, key: "output", value: item.output });
      }
    }
  }
  const messages = requestJson.messages;
  if (Array.isArray(messages)) {
    for (const message of messages) {
      if (
        isJsonObject(message) &&
        (tryGetString(message, "role") ?? "").toLowerCase() === "tool" &&
        typeof message.content === "string"
      ) {
        results.push({
          owner: message,
          key: "content",
          value: message.content,
        });
      }
    }
  }
  return results;
}

function truncateToolResult(value: string, maxChars: number): string {
  const removed = value.length - maxChars;
  const marker = `\n\n[... ${removed} characters omitted from oversized tool result ...]\n\n`;
  const retained = Math.max(0, maxChars - marker.length);
  const headChars = Math.ceil(retained / 2);
  const tailChars = retained - headChars;
  return (
    value.slice(0, headChars) +
    marker +
    value.slice(value.length - tailChars)
  );
}

function shouldRequireInitialLocalToolCall(
  requestJson: JsonObject,
  tooling: OpenAiTooling,
): boolean {
  if (
    tooling.tools.length === 0 ||
    tooling.toolChoiceMode === ToolChoiceModes.None ||
    tooling.toolChoiceMode === ToolChoiceModes.Required ||
    tooling.toolChoiceMode === ToolChoiceModes.Function
  ) {
    return false;
  }
  if (hasPriorToolResult(requestJson)) {
    return false;
  }

  const latestUserText = extractLatestUserText(requestJson).toLowerCase();
  if (!latestUserText.trim()) {
    return false;
  }
  const hasLocalAction = /\b(local shell\/file tools|local shell|file tools|read|inspect|write|create|edit|modify|delete|verify|copy|place|put|clipboard|pbcopy|cat|list|run|execute|pwd|commit|stage|git)\b/.test(
    latestUserText,
  );
  const hasLocalTarget = /\b(file|files|workspace|shell|repo|repository|directory|working directory|cwd|pwd|git|commit|clipboard|previous output|previous response|previous assistant response|docs?|releases?|markdown|md|txt)\b|\.[a-z0-9]{1,8}\b/.test(
    latestUserText,
  );
  return hasLocalAction && hasLocalTarget;
}

function requireLocalToolingIfNeeded(
  requestJson: JsonObject,
  tooling: OpenAiTooling,
): OpenAiTooling {
  if (!shouldRequireInitialLocalToolCall(requestJson, tooling)) {
    return tooling;
  }
  return {
    ...tooling,
    toolChoiceMode: ToolChoiceModes.Required,
    toolChoiceFunctionName: null,
    requiredByLocalAction: true,
  };
}

function buildMappedPromptText(
  userPrompt: string,
  tooling: OpenAiTooling,
): string {
  if (
    tooling.tools.length === 0
  ) {
    return userPrompt;
  }

  const lines = [
    "You are responding through a local Codex harness with shell and file tools.",
    "The local repository is accessible through those tools; do not answer from an M365, Python, or /mnt/file_upload environment.",
    "If tool output is present, treat that output as the source of truth for the local environment.",
  ];
  if (
    tooling.toolChoiceMode === ToolChoiceModes.Required ||
    tooling.toolChoiceMode === ToolChoiceModes.Function
  ) {
    lines.push(
      'A tool call is required for this turn. Respond ONLY as minified JSON with this exact shape: {"tool_calls":[{"name":"<tool-name>","arguments":{}}]}.',
      "Do not include markdown, prose, or a natural-language answer before the tool call.",
      `Available tool names: ${tooling.tools.map((tool) => tool.name).join(", ")}`,
    );
    if (
      tooling.toolChoiceMode === ToolChoiceModes.Function &&
      tooling.toolChoiceFunctionName
    ) {
      lines.push(`Required tool name: ${tooling.toolChoiceFunctionName}`);
    }
  } else {
    lines.push(
      "Use a tool when local repo state is needed; otherwise answer normally using any prior tool outputs.",
    );
  }
  lines.push("User request:", userPrompt);
  return lines.join("\n");
}

export function hasPriorToolResult(requestJson: JsonObject): boolean {
  const input = requestJson.input;
  if (Array.isArray(input)) {
    for (const item of input) {
      if (!isJsonObject(item)) {
        continue;
      }
      const type = (tryGetString(item, "type") ?? "").toLowerCase();
      if (
        type === "function_call_output" ||
        type === "custom_tool_call_output"
      ) {
        return true;
      }
    }
  }

  const messages = requestJson.messages;
  if (Array.isArray(messages)) {
    for (const message of messages) {
      if (!isJsonObject(message)) {
        continue;
      }
      const role = (tryGetString(message, "role") ?? "").toLowerCase();
      if (role === "tool") {
        return true;
      }
    }
  }
  return false;
}

function extractLatestUserText(requestJson: JsonObject): string {
  const input = requestJson.input;
  if (typeof input === "string") {
    return input;
  }
  if (Array.isArray(input)) {
    for (let i = input.length - 1; i >= 0; i--) {
      const item = input[i];
      if (typeof item === "string" && item.trim()) {
        return item;
      }
      if (!isJsonObject(item)) {
        continue;
      }
      const type = (tryGetString(item, "type") ?? "message").toLowerCase();
      const role = (tryGetString(item, "role") ?? "user").toLowerCase();
      if (type === "message" && role === "user") {
        return extractMessageContent(item.content);
      }
    }
  }

  const messages = requestJson.messages;
  if (Array.isArray(messages)) {
    for (let i = messages.length - 1; i >= 0; i--) {
      const message = messages[i];
      if (!isJsonObject(message)) {
        continue;
      }
      const role = (tryGetString(message, "role") ?? "user").toLowerCase();
      if (role === "user") {
        return extractMessageContent(message.content);
      }
    }
  }
  return "";
}

function hasNativeWebSearchIntent(requestJson: JsonObject): boolean {
  const text = extractLatestUserText(requestJson).toLowerCase();
  return (
    /\b(?:search|browse|look up|web|internet|online)\b/.test(text) ||
    /\b(?:latest|current|today|right now|news|headline|recent release)\b/.test(
      text,
    )
  );
}

function resolvePrompt(
  messages: {
    role: string;
    content: string;
    index: number;
    hasImage: boolean;
  }[],
): {
  role: string;
  content: string;
  index: number;
  hasImage: boolean;
} | null {
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === "user") {
      return messages[i];
    }
  }
  return messages.at(-1) ?? null;
}

function normalizeResponsesInput(inputNode: JsonValue | undefined): {
  messages: JsonObject[];
  inputItemsForStorage: JsonValue[];
} {
  if (typeof inputNode === "string" && inputNode.trim()) {
    return {
      messages: [{ role: "user", content: inputNode.trim() }],
      inputItemsForStorage: [
        {
          type: "message",
          role: "user",
          content: [{ type: "input_text", text: inputNode.trim() }],
        },
      ],
    };
  }

  const sourceItems: JsonValue[] = Array.isArray(inputNode)
    ? inputNode
    : inputNode !== undefined && inputNode !== null
      ? [inputNode]
      : [];
  const messages: JsonObject[] = [];
  const inputItemsForStorage: JsonValue[] = [];

  for (const item of sourceItems) {
    if (typeof item === "string") {
      if (!item.trim()) {
        continue;
      }
      messages.push({ role: "user", content: item.trim() });
      inputItemsForStorage.push({
        type: "message",
        role: "user",
        content: [{ type: "input_text", text: item.trim() }],
      });
      continue;
    }

    if (!isJsonObject(item)) {
      continue;
    }

    inputItemsForStorage.push(cloneJsonValue(item));
    const type = (tryGetString(item, "type") ?? "message").toLowerCase();
    if (type === "message") {
      const role = (tryGetString(item, "role") ?? "user").toLowerCase();
      const message: JsonObject = { role };
      if (item.content !== undefined) {
        message.content = cloneJsonValue(item.content);
      } else {
        const text =
          tryGetString(item, "text") ??
          tryGetString(item, "input_text") ??
          tryGetString(item, "output_text");
        if (text) {
          message.content = text;
        }
      }
      messages.push(message);
      continue;
    }

    if (type === "function_call_output" || type === "custom_tool_call_output") {
      const message: JsonObject = {
        role: "tool",
        content: stringifyJsonValue(item.output),
      };
      const toolName = tryGetString(item, "name");
      if (toolName) {
        message.name = toolName;
      }
      const toolCallId =
        tryGetString(item, "call_id") ?? tryGetString(item, "tool_call_id");
      if (toolCallId) {
        message.tool_call_id = toolCallId;
      }
      messages.push(message);
      continue;
    }

    if (type === "function_call" || type === "custom_tool_call") {
      const functionName = tryGetString(item, "name");
      if (!functionName) {
        continue;
      }
      const functionArguments = type === "custom_tool_call"
        ? (tryGetString(item, "input") ?? "")
        : (normalizeFunctionArguments(item.arguments) ?? "{}");
      const toolCall: JsonObject = {
        id:
          tryGetString(item, "call_id") ??
          tryGetString(item, "id") ??
          `call_${messages.length + 1}`,
        type: "function",
        function: {
          name: functionName,
          arguments: functionArguments,
        },
      };
      messages.push({
        role: "assistant",
        content: null,
        tool_calls: [toolCall],
      });
    }
  }

  return { messages, inputItemsForStorage };
}

function mapResponsesTextFormat(requestJson: JsonObject): JsonObject | null {
  const text = requestJson.text;
  if (!isJsonObject(text) || !isJsonObject(text.format)) {
    return null;
  }
  const format = text.format;
  const type = tryGetString(format, "type");
  if (!type) {
    return null;
  }
  const normalizedType = type.toLowerCase();
  if (normalizedType === ResponseFormatTypes.JsonObject) {
    return { type: ResponseFormatTypes.JsonObject };
  }
  if (normalizedType !== ResponseFormatTypes.JsonSchema) {
    return null;
  }

  if (isJsonObject(format.json_schema)) {
    return {
      type: ResponseFormatTypes.JsonSchema,
      json_schema: cloneJsonValue(format.json_schema),
    };
  }

  const jsonSchema: JsonObject = {};
  const name = tryGetString(format, "name");
  if (name) {
    jsonSchema.name = name;
  }
  if (isJsonObject(format.schema)) {
    jsonSchema.schema = cloneJsonValue(format.schema);
  }

  return {
    type: ResponseFormatTypes.JsonSchema,
    json_schema: jsonSchema,
  };
}

function mapResponsesReasoningEffort(requestJson: JsonObject): string | null {
  const reasoning = requestJson.reasoning;
  if (!isJsonObject(reasoning)) {
    return null;
  }
  return tryGetString(reasoning, "effort");
}

function stringifyJsonValue(value: JsonValue | undefined): string {
  if (typeof value === "string") {
    const normalized = normalizeJsonLikeString(value);
    return normalized ?? value;
  }
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value);
}

function normalizeFunctionArguments(value: JsonValue | undefined): string | null {
  if (value === undefined || value === null) {
    return null;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return "{}";
    }
    const parsed = tryParseJsonValueFromText(trimmed);
    return parsed !== null ? JSON.stringify(parsed) : trimmed;
  }
  return JSON.stringify(value);
}

function extractMessageContent(contentNode: JsonValue | undefined): string {
  if (contentNode === undefined || contentNode === null) {
    return "";
  }
  if (typeof contentNode === "string") {
    const normalized = normalizeJsonLikeString(contentNode);
    return normalized ?? contentNode;
  }
  if (isJsonObject(contentNode)) {
    const directText =
      tryGetString(contentNode, "text") ?? tryGetString(contentNode, "value");
    if (directText) {
      const normalized = normalizeJsonLikeString(directText);
      return normalized ?? directText;
    }
    return "";
  }

  if (!Array.isArray(contentNode)) {
    return "";
  }

  const textParts: string[] = [];
  for (const part of contentNode) {
    if (typeof part === "string" && part.trim()) {
      const normalized = normalizeJsonLikeString(part.trim());
      textParts.push(normalized ?? part.trim());
      continue;
    }
    if (!isJsonObject(part)) {
      continue;
    }
    const type = tryGetString(part, "type");
    const isTextPart =
      !type ||
      type.toLowerCase() === "text" ||
      type.toLowerCase() === "input_text";

    if (isTextPart) {
      let partText = tryGetString(part, "text");
      const nestedText = part.text;
      if (!partText && isJsonObject(nestedText)) {
        partText = tryGetString(nestedText, "value");
      }
      if (partText) {
        const normalized = normalizeJsonLikeString(partText.trim());
        textParts.push(normalized ?? partText.trim());
        continue;
      }
    }
  }

  // Images from historical/context messages are intentionally dropped here:
  // only the current user turn is uploaded and annotated, so history must not
  // pretend that an image remains available to the model.
  return textParts.join("\n");
}

function tryExtractAssistantToolCalls(
  message: JsonObject,
  extractedContent: string,
): JsonObject[] {
  if (Array.isArray(message.tool_calls)) {
    return message.tool_calls.filter(isJsonObject);
  }
  if (!extractedContent.trim()) {
    return [];
  }

  const parsedContent = tryParseJsonValueFromText(extractedContent);
  if (parsedContent === null) {
    return [];
  }
  return extractToolCallsFromJsonNode(parsedContent);
}

function extractToolCallsFromJsonNode(node: JsonValue): JsonObject[] {
  if (Array.isArray(node)) {
    return node.filter(isJsonObject).filter(isToolCallLikeObject);
  }
  if (!isJsonObject(node)) {
    return [];
  }

  if (Array.isArray(node.tool_calls)) {
    return node.tool_calls.filter(isJsonObject).filter(isToolCallLikeObject);
  }

  const wrappedFromMessage = node.message;
  if (isJsonObject(wrappedFromMessage) && Array.isArray(wrappedFromMessage.tool_calls)) {
    return wrappedFromMessage.tool_calls
      .filter(isJsonObject)
      .filter(isToolCallLikeObject);
  }

  const wrappedFromChoices = extractToolCallsFromChatCompletionNode(node);
  if (wrappedFromChoices.length > 0) {
    return wrappedFromChoices;
  }

  const wrappedFromOutput = extractToolCallsFromResponsesOutputNode(node);
  if (wrappedFromOutput.length > 0) {
    return wrappedFromOutput;
  }

  if (tryGetString(node, "name")) {
    return [node];
  }

  const functionObject = node.function;
  if (isJsonObject(functionObject) && tryGetString(functionObject, "name")) {
    return [node];
  }

  return [];
}

function isToolCallLikeObject(value: JsonObject): boolean {
  const functionObject = isJsonObject(value.function) ? value.function : null;
  const type = (tryGetString(value, "type") ?? "").toLowerCase();
  return Boolean(
    tryGetString(value, "name") ||
      tryGetString(value, "tool_name") ||
      tryGetString(functionObject, "name") ||
      type === "function" ||
      type === "function_call",
  );
}

function extractToolCallsFromChatCompletionNode(node: JsonObject): JsonObject[] {
  const choices = node.choices;
  if (!Array.isArray(choices)) {
    return [];
  }
  for (const choice of choices) {
    if (!isJsonObject(choice)) {
      continue;
    }
    const message = choice.message;
    if (isJsonObject(message) && Array.isArray(message.tool_calls)) {
      return message.tool_calls.filter(isJsonObject);
    }
    const delta = choice.delta;
    if (isJsonObject(delta) && Array.isArray(delta.tool_calls)) {
      return delta.tool_calls.filter(isJsonObject);
    }
  }
  return [];
}

function extractToolCallsFromResponsesOutputNode(node: JsonObject): JsonObject[] {
  const output = node.output;
  if (!Array.isArray(output)) {
    return [];
  }

  const toolCalls: JsonObject[] = [];
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type !== "function_call" && type !== "custom_tool_call") {
      continue;
    }
    const name = tryGetString(item, "name");
    if (!name) {
      continue;
    }
    if (type === "custom_tool_call") {
      toolCalls.push({
        id:
          tryGetString(item, "call_id") ??
          tryGetString(item, "tool_call_id") ??
          tryGetString(item, "id") ??
          `call_${toolCalls.length + 1}`,
        type: "custom",
        function: {
          name,
          arguments:
            normalizeFunctionArguments(item.input) ??
            normalizeFunctionArguments(item.arguments) ??
            "{}",
        },
      });
    } else {
      toolCalls.push({
        id:
          tryGetString(item, "call_id") ??
          tryGetString(item, "tool_call_id") ??
          tryGetString(item, "id") ??
          `call_${toolCalls.length + 1}`,
        type: "function",
        function: {
          name,
          arguments: normalizeFunctionArguments(item.arguments) ?? "{}",
        },
      });
    }
  }
  return toolCalls;
}

function parseTooling(requestJson: JsonObject): OpenAiTooling {
  const tools: OpenAiToolDefinition[] = [];
  for (const toolsArray of collectResponseToolArrays(requestJson)) {
    for (const toolNode of toolsArray) {
      if (!isJsonObject(toolNode)) {
        continue;
      }
      const type = tryGetString(toolNode, "type")?.toLowerCase();
      if (type !== "function" && type !== "custom") {
        continue;
      }
      const functionObject = isJsonObject(toolNode.function)
        ? toolNode.function
        : toolNode;
      const name = tryGetString(functionObject, "name");
      if (!name) {
        continue;
      }
      const parameters = isJsonObject(functionObject.parameters)
        ? cloneJsonValue(functionObject.parameters)
        : {};

      tools.push({
        name: name.trim(),
        type,
        description: tryGetString(functionObject, "description"),
        parameters,
        format: isJsonObject(functionObject.format)
          ? cloneJsonValue(functionObject.format)
          : null,
      });
    }
  }

  let toolChoiceMode: string =
    tools.length === 0 ? ToolChoiceModes.None : ToolChoiceModes.Auto;
  let toolChoiceFunctionName: string | null = null;
  const toolChoice = requestJson.tool_choice;
  if (typeof toolChoice === "string") {
    const normalized = toolChoice.trim().toLowerCase();
    if (
      normalized === ToolChoiceModes.Auto ||
      normalized === ToolChoiceModes.None ||
      normalized === ToolChoiceModes.Required
    ) {
      toolChoiceMode = normalized;
    }
  } else if (isJsonObject(toolChoice)) {
    const type = tryGetString(toolChoice, "type");
    const functionObject = toolChoice.function;
    if (type?.toLowerCase() === "function" && isJsonObject(functionObject)) {
      toolChoiceFunctionName = tryGetString(functionObject, "name");
      if (toolChoiceFunctionName) {
        toolChoiceMode = ToolChoiceModes.Function;
      }
    }
  }

  return {
    tools,
    toolChoiceMode,
    toolChoiceFunctionName,
    parallelToolCalls:
      tryGetBoolean(requestJson, "parallel_tool_calls") !== false,
    requiredByLocalAction: false,
  };
}

// Codex CLI places its actual tool declarations inside a developer input item
// (`type: additional_tools`), not only at the Responses request top level.
// Keep namespace-contained function declarations too; they are client-offered
// tools and must remain available to strict validation.
function collectResponseToolArrays(requestJson: JsonObject): JsonValue[][] {
  const arrays: JsonValue[][] = [];
  if (Array.isArray(requestJson.tools)) {
    arrays.push(requestJson.tools);
  }
  const input = requestJson.input;
  if (!Array.isArray(input)) {
    return arrays;
  }
  for (const item of input) {
    if (!isJsonObject(item) || tryGetString(item, "type")?.toLowerCase() !== "additional_tools") {
      continue;
    }
    if (Array.isArray(item.tools)) {
      arrays.push(item.tools);
      for (const tool of item.tools) {
        if (isJsonObject(tool) && tryGetString(tool, "type")?.toLowerCase() === "namespace" && Array.isArray(tool.tools)) {
          arrays.push(tool.tools);
        }
      }
    }
  }
  return arrays;
}

function parseResponseFormat(
  requestJson: JsonObject,
): OpenAiResponseFormat | null {
  const responseFormat = requestJson.response_format;
  if (!isJsonObject(responseFormat)) {
    return null;
  }
  const type = tryGetString(responseFormat, "type");
  if (!type) {
    return null;
  }
  const normalizedType = type.toLowerCase();
  if (normalizedType === ResponseFormatTypes.JsonObject) {
    return {
      type: ResponseFormatTypes.JsonObject,
      name: null,
      jsonSchema: null,
    };
  }
  if (normalizedType !== ResponseFormatTypes.JsonSchema) {
    return null;
  }

  const schemaNode = responseFormat.json_schema;
  if (!isJsonObject(schemaNode)) {
    return {
      type: ResponseFormatTypes.JsonSchema,
      name: null,
      jsonSchema: null,
    };
  }

  return {
    type: ResponseFormatTypes.JsonSchema,
    name: tryGetString(schemaNode, "name"),
    jsonSchema: isJsonObject(schemaNode.schema)
      ? cloneJsonValue(schemaNode.schema)
      : null,
  };
}

function buildAdditionalContext(
  messages: { role: string; content: string; index: number }[],
  promptMessageIndex: number,
  requestJson: JsonObject,
  maxContextMessages: number,
  tooling: OpenAiTooling,
  responseFormat: OpenAiResponseFormat | null,
  reasoningEffort: string | null,
  temperature: number | null,
): ContextMessage[] {
  let context: ContextMessage[] = [];
  for (const message of messages) {
    if (message.index === promptMessageIndex || !message.content.trim()) {
      continue;
    }
    context.push({
      text: `${message.role}: ${message.content}`,
      description: null,
    });
  }

  appendCustomContext(context, requestJson.m365_additional_context);

  const systemPrompt = tryGetString(requestJson, "m365_system_prompt");
  if (systemPrompt) {
    context.push({ text: systemPrompt, description: "System prompt override" });
  }

  appendOpenAiCompatibilityContext(
    context,
    tooling,
    responseFormat,
    reasoningEffort,
    temperature,
  );

  if (maxContextMessages > 0 && context.length > maxContextMessages) {
    context = context.slice(context.length - maxContextMessages);
  }
  return context;
}

function appendCustomContext(
  context: ContextMessage[],
  customNode: JsonValue | undefined,
): void {
  if (customNode === undefined || customNode === null) {
    return;
  }
  if (typeof customNode === "string" && customNode.trim()) {
    context.push({ text: customNode.trim(), description: null });
    return;
  }
  if (isJsonObject(customNode)) {
    appendCustomContextObject(context, customNode);
    return;
  }
  if (!Array.isArray(customNode)) {
    return;
  }
  for (const item of customNode) {
    if (typeof item === "string" && item.trim()) {
      context.push({ text: item.trim(), description: null });
      continue;
    }
    if (isJsonObject(item)) {
      appendCustomContextObject(context, item);
    }
  }
}

function appendCustomContextObject(
  context: ContextMessage[],
  contextObject: JsonObject,
): void {
  const text =
    tryGetString(contextObject, "text") ??
    tryGetString(contextObject, "content");
  if (!text) {
    return;
  }
  context.push({
    text: text.trim(),
    description: tryGetString(contextObject, "description"),
  });
}

function appendOpenAiCompatibilityContext(
  context: ContextMessage[],
  tooling: OpenAiTooling,
  responseFormat: OpenAiResponseFormat | null,
  reasoningEffort: string | null,
  temperature: number | null,
): void {
  if (tooling.tools.length > 0) {
    context.push({
      text: [
        "You are connected to a local Codex harness with shell and file tools.",
        "The user's workspace is the local repository described by the Codex turn context, not any M365, Python, or /mnt/file_upload environment you may infer.",
        "For requests that ask you to inspect, edit, verify, or commit local repo files, call the appropriate tool instead of saying the filesystem is unavailable.",
        "If earlier conversation text claimed the local repo was not mounted, treat that claim as stale and recover by using the available tools.",
      ].join(" "),
      description: "Local Codex tool availability",
    });
    context.push({
      text: 'If you call a tool, respond ONLY as minified JSON with this exact shape: {"tool_calls":[{"name":"<tool-name>","arguments":{}}]}. No markdown, no prose, no extra keys.',
      description: "OpenAI tool-calling contract",
    });
    context.push({
      text: "If no tool call is needed, return a normal assistant response (unless response_format requires JSON-only output).",
      description: "OpenAI tool-calling contract",
    });
    context.push({
      text: JSON.stringify(
        tooling.tools.map((tool) => ({
          name: tool.name,
          type: tool.type,
          description: tool.description,
          parameters: cloneJsonValue(tool.parameters),
          format: tool.format ? cloneJsonValue(tool.format) : undefined,
        })),
      ),
      description: "Available tools",
    });

    if (tooling.toolChoiceMode === ToolChoiceModes.None) {
      context.push({
        text: "Tool calls are disabled for this response.",
        description: "Tool choice",
      });
    } else if (tooling.toolChoiceMode === ToolChoiceModes.Required) {
      context.push({
        text: "A tool call is required in this turn. Do not return plain assistant text first.",
        description: "Tool choice",
      });
    } else if (
      tooling.toolChoiceMode === ToolChoiceModes.Function &&
      tooling.toolChoiceFunctionName
    ) {
      context.push({
        text: `You must call only tool '${tooling.toolChoiceFunctionName}' in this turn. Do not return plain assistant text first.`,
        description: "Tool choice",
      });
    }
  }

  if (responseFormat) {
    if (responseFormat.type === ResponseFormatTypes.JsonObject) {
      context.push({
        text: "Return ONLY a valid JSON object and no markdown.",
        description: "OpenAI response_format",
      });
    } else if (responseFormat.type === ResponseFormatTypes.JsonSchema) {
      context.push({
        text: "Return ONLY valid JSON that conforms to the provided JSON schema.",
        description: "OpenAI response_format",
      });
      if (responseFormat.jsonSchema) {
        context.push({
          text: JSON.stringify(responseFormat.jsonSchema),
          description: "JSON schema",
        });
      }
    }
  }

  if (reasoningEffort) {
    context.push({
      text: `Reasoning effort preference: ${reasoningEffort}.`,
      description: "Reasoning hint",
    });
  }
  if (temperature !== null) {
    context.push({
      text: `Sampling temperature preference: ${temperature}.`,
      description: "Generation hint",
    });
  }
}

function normalizeJsonLikeString(rawText: string): string | null {
  const parsed = tryParseJsonValueFromText(rawText);
  if (parsed === null) {
    return null;
  }
  return JSON.stringify(parsed);
}

function tryParseJsonValueFromText(rawText: string): JsonValue | null {
  for (const candidate of enumerateJsonCandidates(rawText)) {
    const parsed = tryParseJsonValue(candidate);
    if (parsed !== null) {
      return parsed;
    }
  }
  return null;
}

function tryParseJsonValue(rawText: string): JsonValue | null {
  if (!rawText.trim()) {
    return null;
  }
  try {
    return JSON.parse(rawText) as JsonValue;
  } catch {
    return null;
  }
}

function* enumerateJsonCandidates(rawText: string): Iterable<string> {
  const trimmed = rawText.trim();
  if (!trimmed) {
    return;
  }
  const seen = new Set<string>();
  const yieldCandidate = function* (candidate: string): Iterable<string> {
    const normalized = candidate.trim();
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    yield normalized;
  };

  yield* yieldCandidate(trimmed);

  let cursor = 0;
  while (cursor < rawText.length) {
    const fenceStart = rawText.indexOf("```", cursor);
    if (fenceStart < 0) {
      break;
    }
    const bodyStart = rawText.indexOf("\n", fenceStart + 3);
    if (bodyStart < 0) {
      break;
    }
    const fenceEnd = rawText.indexOf("```", bodyStart + 1);
    if (fenceEnd < 0) {
      break;
    }
    const body = rawText.slice(bodyStart + 1, fenceEnd).trim();
    if (body) {
      yield* yieldCandidate(body);
    }
    cursor = fenceEnd + 3;
  }

  for (const balanced of extractBalancedJsonSegments(rawText)) {
    yield* yieldCandidate(balanced);
  }
}

function* extractBalancedJsonSegments(rawText: string): Iterable<string> {
  const maxCandidates = 128;
  let emitted = 0;
  for (let start = 0; start < rawText.length; start++) {
    if (emitted >= maxCandidates) {
      break;
    }
    const opening = rawText[start];
    if (opening !== "{" && opening !== "[") {
      continue;
    }
    const balanced = extractBalancedJsonSegment(rawText, start, opening);
    if (!balanced) {
      continue;
    }
    emitted++;
    yield balanced;
  }
}

function extractBalancedJsonSegment(
  rawText: string,
  start: number,
  opening: string,
): string | null {
  const closing = opening === "{" ? "}" : "]";
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = start; index < rawText.length; index++) {
    const ch = rawText[index];
    if (!ch) {
      continue;
    }
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === "\\") {
        escaped = true;
        continue;
      }
      if (ch === "\"") {
        inString = false;
      }
      continue;
    }

    if (ch === "\"") {
      inString = true;
      continue;
    }
    if (ch === opening) {
      depth++;
      continue;
    }
    if (ch === closing) {
      depth--;
      if (depth === 0) {
        return rawText.slice(start, index + 1).trim();
      }
    }
  }
  return null;
}

function hasImageContent(contentNode: JsonValue | undefined): boolean {
  if (isImageContentPart(contentNode)) {
    return true;
  }
  return Array.isArray(contentNode) && contentNode.some(isImageContentPart);
}

function isImageContentPart(value: JsonValue | undefined): value is JsonObject {
  if (!isJsonObject(value)) {
    return false;
  }
  const type = (tryGetString(value, "type") ?? "").trim().toLowerCase();
  // Must stay in sync with `isImagePart` in ./image-input. A bare `url` field
  // is deliberately excluded: this predicate also drives prompt sanitization,
  // and matching on `url` alone would strip unrelated request nodes such as
  // tool schemas that happen to declare a `url` property.
  return (
    type === "input_image" ||
    type === "image_url" ||
    type === "image" ||
    value.image_url !== undefined
  );
}

function parseCurrentTurnImages(
  requestJson: JsonObject,
  options: WrapperOptions,
) {
  return parseImageInputs(
    extractCurrentUserContent(requestJson),
    imageInputLimits(options),
  );
}

function imageInputLimits(options: WrapperOptions): {
  enabled?: boolean;
  maxImages?: number;
  maxBytesPerImage?: number;
  maxTotalBytes?: number;
  allowedMimeTypes?: string[];
} {
  return {
    enabled: options.substrate.imageUploadEnabled,
    maxImages: options.substrate.maxImagesPerRequest,
    maxBytesPerImage: options.substrate.maxBytesPerImage,
    maxTotalBytes: options.substrate.maxTotalImageBytes,
    allowedMimeTypes: options.substrate.allowedImageMimeTypes,
  };
}

function extractCurrentUserContent(
  requestJson: JsonObject,
): JsonValue | undefined {
  const messages = requestJson.messages;
  if (Array.isArray(messages)) {
    for (let index = messages.length - 1; index >= 0; index--) {
      const message = messages[index];
      if (
        isJsonObject(message) &&
        (tryGetString(message, "role") ?? "user").toLowerCase() === "user"
      ) {
        return message.content;
      }
    }
  }

  const input = requestJson.input;
  if (Array.isArray(input)) {
    for (let index = input.length - 1; index >= 0; index--) {
      const item = input[index];
      if (!isJsonObject(item)) {
        continue;
      }
      const type = (tryGetString(item, "type") ?? "message").toLowerCase();
      const role = (tryGetString(item, "role") ?? "user").toLowerCase();
      if (type === "message" && role === "user") {
        return item.content;
      }
    }
  }
  return undefined;
}

function convertToolCallsToContextText(toolCalls: JsonValue[]): string {
  const calls = toolCalls.filter(isJsonObject);
  return calls.length === 0
    ? ""
    : `assistant tool_calls: ${JSON.stringify(calls)}`;
}

function buildLocationHint(
  requestJson: JsonObject,
  defaultTimeZone: string,
): JsonObject {
  let locationHint: JsonObject = {};
  const explicitLocationHint = requestJson.m365_location_hint;
  if (isJsonObject(explicitLocationHint)) {
    locationHint = cloneJsonValue(explicitLocationHint);
  } else if (
    typeof explicitLocationHint === "string" &&
    explicitLocationHint.trim()
  ) {
    locationHint.timeZone = explicitLocationHint.trim();
  }

  const timeZoneOverride = tryGetString(requestJson, "m365_time_zone");
  if (timeZoneOverride) {
    locationHint.timeZone = timeZoneOverride;
  }

  if (!tryGetString(locationHint, "timeZone")) {
    locationHint.timeZone = defaultTimeZone?.trim()
      ? defaultTimeZone
      : "America/New_York";
  }

  const countryOrRegion = tryGetString(requestJson, "m365_country_or_region");
  if (countryOrRegion && !tryGetString(locationHint, "countryOrRegion")) {
    locationHint.countryOrRegion = countryOrRegion;
  }
  return locationHint;
}

function buildContextualResources(requestJson: JsonObject): JsonObject | null {
  const custom = requestJson.m365_contextual_resources;
  if (isJsonObject(custom)) {
    return cloneJsonValue(custom);
  }
  const direct = requestJson.contextualResources;
  if (isJsonObject(direct)) {
    return cloneJsonValue(direct);
  }
  return null;
}
