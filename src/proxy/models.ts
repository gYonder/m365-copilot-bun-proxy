import type { JsonObject } from "./types";

// Codex M365 has one supported route. Keep the advertised catalog narrow so
// clients cannot select unverified M365 selectors.
export const AvailableModelIds = ["gpt-5.6-sol"] as const;
export const ConfiguredContextLimit = 400_000;

const CodexBaseInstructions =
  "You are Codex, a coding agent running in a local CLI. You and the user share the same workspace. Use the available shell/file tools to inspect and change files when a task requires local state, and verify the result before claiming it is done.";

export function buildModelsResponse(): JsonObject {
  return {
    object: "list",
    data: AvailableModelIds.map((id) => ({
      id,
      object: "model",
      created: 0,
      owned_by: "microsoft-365-copilot",
    })),
    models: AvailableModelIds.map((id, index) => buildCodexModelInfo(id, index)),
  };
}

function buildCodexModelInfo(
  id: (typeof AvailableModelIds)[number],
  index: number,
): JsonObject {
  // M365 exposes one verified GPT-5.6 reasoning selector. All Codex effort
  // levels map to that selector; retaining the standard levels keeps
  // continuation requests valid without inventing unobserved M365 tones.
  const supportedReasoningLevels = [
    { effort: "none", description: "No extra reasoning" },
    { effort: "low", description: "Light reasoning" },
    { effort: "medium", description: "Balanced reasoning" },
    { effort: "high", description: "Deep reasoning" },
    { effort: "xhigh", description: "Very deep reasoning" },
    { effort: "max", description: "Maximum reasoning" },
  ];

  return {
    slug: id,
    display_name: id,
    description: "Microsoft 365 Copilot via the local Bun proxy",
    default_reasoning_level: "high",
    supported_reasoning_levels: supportedReasoningLevels,
    shell_type: "shell_command",
    visibility: "list",
    supported_in_api: true,
    priority: AvailableModelIds.length - index,
    upgrade: null,
    base_instructions: CodexBaseInstructions,
    supports_reasoning_summaries: false,
    support_verbosity: true,
    default_verbosity: "low",
    apply_patch_tool_type: "freeform",
    truncation_policy: { mode: "tokens", limit: 10_000 },
    supports_parallel_tool_calls: true,
    supports_image_detail_original: false,
    context_window: ConfiguredContextLimit,
    experimental_supported_tools: [],
    input_modalities: ["text"],
  };
}
