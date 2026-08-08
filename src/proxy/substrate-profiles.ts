import { canonicalizeJson, hashCanonicalJson } from "./canonical-json";
import { resolveSubstrateCapabilities } from "./substrate-capabilities";
import type {
  JsonObject,
  OpenAiToolDefinition,
  ParsedOpenAiRequest,
  WrapperOptions,
} from "./types";
import { isJsonObject, tryGetString } from "./utils";

export type RequestProfile = {
  profileId: string;
  compatibilityKey: string;
  taskScope: string;
  carrierModel: string;
  upstreamSelector: string;
  localTools: boolean;
  hostedWebSearch: boolean;
  nativeCodeInterpreter: boolean;
  imageInput: boolean;
  localToolProtocol: string;
  localToolSchemaHash: string;
  webSearchMode: "router" | "disabled";
  transformMode: string;
  nativeFeatureSetHash: string;
};

export type RequestProfileInput = {
  requestJson: JsonObject;
  request: ParsedOpenAiRequest;
  options: WrapperOptions;
  transport: string;
};

const HostedWebSearchTypes = new Set([
  "web_search",
  "web_search_preview",
  "hosted_web_search",
  "web_search_call",
]);

const NativeCodeInterpreterTypes = new Set([
  "code_interpreter",
  "code_interpreter_preview",
  "code_interpreter_call",
]);

const ImageInputTypes = new Set(["image", "image_url", "input_image"]);

export function deriveRequestProfile({
  requestJson,
  request,
  options,
  transport,
}: RequestProfileInput): RequestProfile {
  const structuredTypes = collectStructuredTypes(requestJson);
  const hostedWebSearch =
    [...structuredTypes].some((type) => HostedWebSearchTypes.has(type)) ||
    isJsonObject(requestJson.web_search_options);
  const nativeCodeInterpreter = [...structuredTypes].some((type) =>
    NativeCodeInterpreterTypes.has(type),
  );
  const imageInput =
    [...structuredTypes].some((type) => ImageInputTypes.has(type)) ||
    hasStructuredKey(requestJson, "image_url");

  const localTools = request.tooling.tools.length > 0;
  const localToolSchemaHash = hashCanonicalJson(
    request.tooling.tools
      .map(toCanonicalTool)
      .sort((left, right) =>
        canonicalizeJson(left).localeCompare(canonicalizeJson(right)),
      ),
  );
  const localToolProtocol = [...new Set(
    request.tooling.tools.map((tool) => tool.type),
  )]
    .sort()
    .join("+") || "none";
  const capabilities = resolveSubstrateCapabilities(
    request.model,
    options.substrate,
  );
  const nativeFeatureSetHash = hashCanonicalJson({
    semantic_protocol_version: "m365.substrate.profile.v1",
    transport: normalizeValue(transport),
    invocation_target: capabilities.invocationTarget,
    invocation_type: capabilities.invocationType,
    tone: capabilities.tone,
    message_type: capabilities.messageType,
    streaming_mode: capabilities.streamingMode,
    spoken_text_mode: capabilities.spokenTextMode,
    locale: capabilities.locale,
    experience_type: capabilities.experienceType,
    options_sets: [...capabilities.optionsSets].sort(),
    allowed_message_types: [...capabilities.allowedMessageTypes].sort(),
    entity_annotation_types: [...capabilities.entityAnnotationTypes].sort(),
  });

  const taskScope = normalizeTaskScope(options.taskScope);
  const profileId = buildProfileId({
    hostedWebSearch,
    imageInput,
    nativeCodeInterpreter,
  });
  const compatibilityFields = {
    profile_id: profileId,
    task_scope: taskScope,
    carrier_model: normalizeValue(request.model),
    upstream_selector: capabilities.tone,
    local_tools: localTools,
    local_tool_protocol: localToolProtocol,
    local_tool_schema_hash: localToolSchemaHash,
    parallel_tool_calls: request.tooling.parallelToolCalls,
    hosted_web_search: hostedWebSearch,
    native_code_interpreter: nativeCodeInterpreter,
    image_input: imageInput,
    web_search_mode: hostedWebSearch ? "router" : "disabled",
    transform_mode: normalizeValue(request.transformMode),
    reasoning_effort: request.reasoningEffort,
    response_format: request.responseFormat,
    native_feature_set_hash: nativeFeatureSetHash,
  };

  return {
    profileId,
    compatibilityKey: `m365.profile.v1:${hashCanonicalJson(compatibilityFields)}`,
    taskScope,
    carrierModel: request.model,
    upstreamSelector: capabilities.tone,
    localTools,
    hostedWebSearch,
    nativeCodeInterpreter,
    imageInput,
    localToolProtocol,
    localToolSchemaHash,
    webSearchMode: hostedWebSearch ? "router" : "disabled",
    transformMode: request.transformMode,
    nativeFeatureSetHash,
  };
}

export function isSupportedTaskScope(taskScope: string): boolean {
  return normalizeTaskScope(taskScope) === "coding_project";
}

function normalizeTaskScope(taskScope: string | null | undefined): string {
  const normalized = normalizeValue(taskScope);
  return normalized || "coding_project";
}

function normalizeValue(value: string | null | undefined): string {
  return value?.trim().toLowerCase() ?? "";
}

function buildProfileId({
  hostedWebSearch,
  imageInput,
  nativeCodeInterpreter,
}: {
  hostedWebSearch: boolean;
  imageInput: boolean;
  nativeCodeInterpreter: boolean;
}): string {
  const features: string[] = [];
  if (hostedWebSearch) {
    features.push("coding_web_search");
  } else if (imageInput) {
    features.push("image_input");
  } else {
    features.push("coding");
  }
  if (imageInput && !features.includes("image_input")) {
    features.push("image_input");
  }
  if (nativeCodeInterpreter) {
    features.push("code_interpreter");
  }
  return `m365.substrate.${features.join("_")}`;
}

function toCanonicalTool(tool: OpenAiToolDefinition): JsonObject {
  return {
    name: tool.name,
    type: tool.type,
    description: tool.description,
    parameters: tool.parameters,
    format: tool.format,
  };
}

function collectStructuredTypes(value: unknown): Set<string> {
  const types = new Set<string>();
  if (Array.isArray(value)) {
    for (const item of value) {
      for (const type of collectStructuredTypes(item)) {
        types.add(type);
      }
    }
    return types;
  }
  if (!isJsonObject(value)) {
    return types;
  }
  const type = tryGetString(value, "type")?.trim().toLowerCase();
  if (type) {
    types.add(type);
  }
  for (const child of Object.values(value)) {
    if (Array.isArray(child) || isJsonObject(child)) {
      for (const nestedType of collectStructuredTypes(child)) {
        types.add(nestedType);
      }
    }
  }
  return types;
}

function hasStructuredKey(value: unknown, key: string): boolean {
  if (Array.isArray(value)) {
    return value.some((item) => hasStructuredKey(item, key));
  }
  if (!isJsonObject(value)) {
    return false;
  }
  if (Object.prototype.hasOwnProperty.call(value, key)) {
    return true;
  }
  return Object.entries(value).some(([childKey, child]) => {
    if (
      childKey === "tools" ||
      childKey === "parameters" ||
      childKey === "properties" ||
      childKey === "format" ||
      childKey === "function"
    ) {
      return false;
    }
    return (
      (Array.isArray(child) || isJsonObject(child)) &&
      hasStructuredKey(child, key)
    );
  });
}
