import type { SubstrateOptions } from "./types";

export const ObservedSubstrateTones = {
  Magic: "magic",
  Chat: "Chat",
  Reasoning: "Reasoning",
  Gpt55Chat: "Gpt_5_5_Chat",
  Gpt56Reasoning: "Gpt_5_6_Reasoning",
} as const;

export const DefaultSubstrateCapabilities = {
  messageType: "Chat",
  streamingMode: "ConciseWithPadding",
  spokenTextMode: "None",
  invocationTarget: "chat",
  invocationType: 4,
  locale: "en-US",
  experienceType: "Default",
} as const;

export type ResolvedSubstrateCapabilities = {
  tone: string;
  messageType: string;
  streamingMode: string;
  spokenTextMode: string;
  invocationTarget: string;
  invocationType: number;
  locale: string;
  experienceType: string;
  optionsSets: string[];
  allowedMessageTypes: string[];
  entityAnnotationTypes: string[];
};

export function resolveSubstrateTone(
  model: string | null | undefined,
): string {
  switch (model?.trim().toLowerCase() ?? "") {
    case "gpt-5.6-sol":
      return ObservedSubstrateTones.Gpt56Reasoning;
    case "m365-copilot-quick":
      return ObservedSubstrateTones.Chat;
    case "m365-copilot-reasoning":
      return ObservedSubstrateTones.Reasoning;
    case "m365-copilot-gpt5.5-quick":
      return ObservedSubstrateTones.Gpt55Chat;
    default:
      return ObservedSubstrateTones.Magic;
  }
}

export function resolveSubstrateCapabilities(
  model: string | null | undefined,
  substrate: SubstrateOptions,
): ResolvedSubstrateCapabilities {
  return {
    tone: resolveSubstrateTone(model),
    messageType: DefaultSubstrateCapabilities.messageType,
    streamingMode: DefaultSubstrateCapabilities.streamingMode,
    spokenTextMode: DefaultSubstrateCapabilities.spokenTextMode,
    invocationTarget:
      substrate.invocationTarget?.trim() ||
      DefaultSubstrateCapabilities.invocationTarget,
    invocationType:
      substrate.invocationType > 0
        ? substrate.invocationType
        : DefaultSubstrateCapabilities.invocationType,
    locale: substrate.locale?.trim() || DefaultSubstrateCapabilities.locale,
    experienceType:
      substrate.experienceType?.trim() ||
      DefaultSubstrateCapabilities.experienceType,
    optionsSets: normalizeList(substrate.optionsSets),
    allowedMessageTypes: normalizeList(substrate.allowedMessageTypes),
    entityAnnotationTypes: normalizeList(substrate.entityAnnotationTypes),
  };
}

function normalizeList(values: string[]): string[] {
  return values.map((value) => value.trim()).filter(Boolean);
}
