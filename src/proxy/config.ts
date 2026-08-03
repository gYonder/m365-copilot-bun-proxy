import { promises as fs } from "node:fs";
import path from "node:path";
import { z } from "zod";
import {
  LogLevels,
  OpenAiTransformModes,
  PlaywrightBrowsers,
  type JsonObject,
  type WrapperOptions,
} from "./types";
import { deepMerge, isJsonObject, parseEnvValue, setDeepValue } from "./utils";

const LogLevelSchema = z.preprocess(
  (value) => (typeof value === "string" ? value.trim().toLowerCase() : value),
  z.enum([
    LogLevels.Trace,
    LogLevels.Debug,
    LogLevels.Info,
    LogLevels.Warning,
    LogLevels.Error,
  ]),
);

const OpenAiTransformModeSchema = z.preprocess(
  (value) => (typeof value === "string" ? value.trim().toLowerCase() : value),
  z.enum([OpenAiTransformModes.Simulated, OpenAiTransformModes.Mapped]),
);

const PlaywrightBrowserSchema = z.preprocess(
  (value) => {
    if (typeof value !== "string") {
      return value;
    }
    const normalized = value.trim().toLowerCase();
    return normalized === "msedge" ? PlaywrightBrowsers.Edge : normalized;
  },
  z.enum([
    PlaywrightBrowsers.Edge,
    PlaywrightBrowsers.Chrome,
    PlaywrightBrowsers.Chromium,
    PlaywrightBrowsers.Firefox,
    PlaywrightBrowsers.Webkit,
  ]),
);

function defaultSubstrateDeviceOS(): string {
  switch (process.platform) {
    case "darwin":
      return "macOS";
    case "win32":
      return "Windows";
    case "linux":
      return "Linux";
    default:
      return process.platform || "Unknown";
  }
}

const WrapperOptionsSchema = z.object({
  listenUrl: z.string().default("http://localhost:4000"),
  debugPath: z.string().nullable().default("./Logs"),
  logLevel: LogLevelSchema.default(LogLevels.Info),
  logStreamingResponseBody: z.boolean().default(false),
  openAiTransformMode: OpenAiTransformModeSchema.default(
    OpenAiTransformModes.Simulated,
  ),
  temporaryChat: z.boolean().default(true),
  ignoreIncomingAuthorizationHeader: z.boolean().default(true),
  playwrightBrowser: PlaywrightBrowserSchema.default(PlaywrightBrowsers.Edge),
  transport: z.string().default("graph"),
  graphBaseUrl: z.string().default("https://graph.microsoft.com"),
  createConversationPath: z.string().default("/beta/copilot/conversations"),
  chatPathTemplate: z
    .string()
    .default("/beta/copilot/conversations/{conversationId}/chat"),
  chatOverStreamPathTemplate: z
    .string()
    .default("/beta/copilot/conversations/{conversationId}/chatOverStream"),
  substrate: z
    .object({
      hubPath: z
        .string()
        .default("wss://substrate.office.com/m365Copilot/Chathub"),
      source: z.string().default("officeweb"),
      quoteSourceInQuery: z.boolean().default(true),
      scenario: z.string().default("OfficeWebIncludedCopilot"),
      origin: z.string().default("https://m365.cloud.microsoft"),
      product: z.string().nullable().default("Office"),
      agentHost: z.string().nullable().default("Bizchat.FullScreen"),
      licenseType: z.string().nullable().default("Starter"),
      isEdu: z.boolean().default(false),
      agent: z.string().nullable().default("web"),
      variants: z
        .string()
        .nullable()
        .default(
          "EnableMcpServerWidgets,feature.EnableMcpServerWidgets,feature.EnableImageGenInsufficientTokensThrottled,feature.EnableImageGenSystemCapacityThrottled,feature.EnableLuForChatCIQ,feature.enableChatCIQPlugin,EnableRequestPlugins,feature.EnableSensitivityLabels,EnableUnsupportedUrlDetector,feature.IsCustomEngineCopilotEnabled,feature.bizchatfluxv3,feature.enablechatpages,feature.enableCodeCanvas,feature.turnOnDARecommendation,feature.IsStreamingModeInChatRequestEnabled,IncludeSourceAttributionsConcise,SkipPublishEmptyMessage,feature.EnableDeduplicatingSourceAttributions,feature.IsCitationsReferencesOutputEnabled,feature.enableDeltaStreamingForReferences,feature.enableIncludeReferencesInDeltaResponse,feature.enablereferencesforagents,feature.EnableCodeInterpreterConversion,agt_module_attr_enableReferencesForCodeInterpreter,agt_module_enableCodeInterpreterHallucinatedUrlFilter,Enable3PActionProgressMessages,feature.EnableCIQDesktopDisplay,feature.enableClientWebRtc,feature.EnableMeetingRecapOfSeriesMeetingWithCiq,feature.EnableReferencesListCompleteSignal,feature.StorageMessageSplitDisabled,feature.EnableCuaTakeControlApi,SingletonEnvOn,cdxenablefccinmainline,EnableComposeWidget,-agt_researcheragent_enableMemoryRead,feature.cwcallowedos,feature.EnableMergingPureDeltas,feature.disabledisallowedmsgs,feature.enableCitationsForSynthesisData,feature.EnableConversationShareApis,feature.enableGenerateGraphicArtOptionsSet,cdximagen,feature.EnableUpdatedUXForConfirmationDialog,feature.EnableContentApiandDocTypeHtmlInRichAnswers,cdxgrounding_api_v2_rich_web_answers_reference_bottom_force,cdxenablerenderforisocomp,feature.EnableClientFileURLSupportForOfficeWebPaidCopilot,feature.EnableDesignEditorImageGrounding,feature.EnableDesignerEditor,feature.EnableSkipRehydrationForSpeCIdImages,feature.EnablePersonalization,rich_responses,feature.EnableBase64DataInMessageAnnotations,feature.EnableSkipEmittingMessageOnFlush,feature.EnableRemoveEmptySourceAttributions,feature.EnableRemoveStreamingMode,feature.OfficeWebToHelix,feature.OfficeDesktopToHelix,feature.M365TeamsHubToHelix,feature.OwaHubToHelix,feature.MonarchHubToHelix,feature.Win32OutlookHubToHelix,feature.MacOutlookHubToHelix,Agt_bizchat_enableGpt5ForHelix",
        ),
      clientPlatform: z.string().default("mcmcopilot-web"),
      clientAppName: z.string().default("Office"),
      clientEntrypoint: z.string().default("mcmcopilot-officeweb"),
      clientAppType: z.string().default("Web"),
      productEntryPoint: z.string().default("ChatPanel"),
      productCategory: z.string().default("Chat"),
      deviceOS: z.string().default(defaultSubstrateDeviceOS()),
      deviceType: z.string().default("Desktop"),
      productThreadType: z.string().default("Office"),
      invocationTimeoutSeconds: z.number().int().default(120),
      handshakeTimeoutSeconds: z.number().int().positive().optional(),
      turnTimeoutSeconds: z.number().int().positive().optional(),
      taskTimeoutSeconds: z.number().int().positive().default(900),
      keepAliveSeconds: z.number().int().default(15),
      optionsSets: z
        .array(z.string())
        .default([
          "search_result_progress_messages_with_search_queries",
          "cwc_flux_image",
          "cwc_code_interpreter",
          "cwc_code_interpreter_amsfix",
          "cwcfluxgptv",
          "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
          "gptvnorm2048",
          "cwc_code_interpreter_citation_fix",
          "code_interpreter_interactive_charts",
          "cwc_code_interpreter_interactive_charts_inline_image",
          "code_interpreter_matplotlib_patching",
          "cwc_fileupload_odb",
          "update_memory_plugin",
          "add_custom_instructions",
          "cwc_flux_v3",
          "flux_v3_progress_messages",
          "enable_batch_token_processing",
          "enable_gg_gpt",
        ]),
      allowedMessageTypes: z
        .array(z.string())
        .default([
          "Chat",
          "Suggestion",
          "InternalSearchQuery",
          "Disengaged",
          "InternalLoaderMessage",
          "Progress",
          "GeneratedCode",
          "RenderCardRequest",
          "AdsQuery",
          "SemanticSerp",
          "GenerateContentQuery",
          "GenerateGraphicArt",
          "SearchQuery",
          "ConfirmationCard",
          "AuthError",
          "DeveloperLogs",
          "TriggerPlugin",
          "HintInvocation",
          "MemoryUpdate",
          "EndOfRequest",
          "TriggerConfirmation",
          "ResumeInvokeAction",
          "ResumeUserInputRequest",
          "TriggerUserInputRequest",
          "EscapeHatch",
          "TriggerPluginAuth",
          "ResumePluginAuth",
          "SideBySide",
          "ReferencesListComplete",
          "SwitchRespondingEndpoint",
        ]),
      invocationTarget: z.string().default("chat"),
      invocationType: z.number().int().default(4),
      locale: z.string().default("en-US"),
      experienceType: z.string().default("Default"),
      entityAnnotationTypes: z
        .array(z.string())
        .default(["People", "File", "Event", "Email", "TeamsMessage"]),
      earlyCompleteOnSimulatedPayload: z.boolean().default(false),
      incrementalSimulatedContentStreaming: z.boolean().default(false),
      concurrencyLimit: z.number().int().min(0).default(2),
      acquireTimeoutMs: z.number().int().min(0).default(0),
      maxSendChars: z.number().int().min(0).default(500_000),
      truncateBeforeSending: z.boolean().default(true),
    })
    .default({}),
  defaultModel: z.string().default("gpt-5.6-sol"),
  defaultTimeZone: z.string().default("America/New_York"),
  conversationTtlMinutes: z.number().int().default(180),
  maxAdditionalContextMessages: z.number().int().default(16),
  includeConversationIdInResponseBody: z.boolean().default(true),
  retrySimulatedToollessResponses: z.boolean().default(true),
  logStdout: z.boolean().default(false),
  confabRetries: z.number().int().min(0).default(1),
  msalAuth: z.boolean().default(true),
  imageGeneration: z
    .object({
      enabled: z.boolean().default(false),
      maxPromptChars: z.number().int().positive().default(4_000),
      maxImages: z.number().int().min(1).max(4).default(1),
      maxArtifactBytes: z.number().int().positive().default(20_000_000),
      timeoutMs: z.number().int().positive().default(120_000),
      concurrencyLimit: z.number().int().min(1).default(1),
      allowedMimeTypes: z
        .array(z.string())
        .min(1)
        .default(["image/png", "image/jpeg", "image/webp"]),
    })
    .default({}),
});

export async function loadWrapperOptions(cwd: string): Promise<WrapperOptions> {
  const rootConfig: JsonObject = {};
  const baseConfig = await readJsonFile(path.join(cwd, "config.json"));
  deepMerge(rootConfig, baseConfig ?? {});

  const env = process.env.NODE_ENV;
  if (env?.trim()) {
    const envConfig = await readJsonFile(path.join(cwd, `config.${env}.json`));
    deepMerge(rootConfig, envConfig ?? {});
  }

  applyConfigEnvOverrides(rootConfig, process.env);
  applyFriendlyEnvOverrides(rootConfig, process.env);
  return normalizeWrapperOptions(rootConfig);
}

function normalizeWrapperOptions(wrapper: JsonObject): WrapperOptions {
  const normalized: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(wrapper)) {
    normalized[key] = value;
  }

  for (const [key, value] of Object.entries(normalized)) {
    if (typeof value !== "string") {
      continue;
    }
    if (/^-?\d+(\.\d+)?$/.test(value.trim())) {
      normalized[key] = Number.parseFloat(value);
      continue;
    }
    const lowered = value.trim().toLowerCase();
    if (lowered === "true" || lowered === "false") {
      normalized[key] = lowered === "true";
    }
  }

  return WrapperOptionsSchema.parse(normalized) as WrapperOptions;
}

function applyConfigEnvOverrides(
  wrapper: JsonObject,
  env: NodeJS.ProcessEnv,
): void {
  for (const [key, value] of Object.entries(env)) {
    if (!value) {
      continue;
    }
    if (!key.toUpperCase().startsWith("CONFIG__")) {
      continue;
    }
    const pathParts = key
      .split("__")
      .slice(1)
      .filter((part) => part.trim().length > 0);
    if (pathParts.length === 0) {
      continue;
    }
    setDeepValue(wrapper, pathParts, parseEnvValue(value));
  }
}

// Convenience aliases for a handful of runtime-robustness knobs so operators can
// flip them without the verbose CONFIG__ path syntax.
function applyFriendlyEnvOverrides(
  wrapper: JsonObject,
  env: NodeJS.ProcessEnv,
): void {
  const isTruthy = (raw: string): boolean =>
    ["1", "true", "yes", "on"].includes(raw.trim().toLowerCase());

  const stdout = env.M365_LOG_STDOUT;
  if (stdout?.trim()) {
    wrapper.logStdout = isTruthy(stdout);
  }

  if (env.M365_NO_CONFAB_RETRY?.trim() && isTruthy(env.M365_NO_CONFAB_RETRY)) {
    wrapper.confabRetries = 0;
  }
  const confabRetries = env.M365_CONFAB_RETRIES;
  if (confabRetries?.trim() && /^\d+$/.test(confabRetries.trim())) {
    wrapper.confabRetries = Number.parseInt(confabRetries.trim(), 10);
  }

  const substrate = isJsonObject(wrapper.substrate)
    ? (wrapper.substrate as JsonObject)
    : (wrapper.substrate = {} as JsonObject);

  const concurrency = env.M365_SUBSTRATE_CONCURRENCY;
  if (concurrency?.trim() && /^\d+$/.test(concurrency.trim())) {
    substrate.concurrencyLimit = Number.parseInt(concurrency.trim(), 10);
  }
  const maxSendChars = env.M365_MAX_SEND_CHARS;
  if (maxSendChars?.trim() && /^\d+$/.test(maxSendChars.trim())) {
    substrate.maxSendChars = Number.parseInt(maxSendChars.trim(), 10);
  }
}

async function readJsonFile(filePath: string): Promise<JsonObject | null> {
  try {
    const content = await fs.readFile(filePath, "utf8");
    const parsed = JSON.parse(content) as unknown;
    return isJsonObject(parsed) ? parsed : null;
  } catch {
    return null;
  }
}
