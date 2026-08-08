import type { JsonObject, JsonValue } from "./types";
import { isJsonObject, tryGetString } from "./utils";

const CitationPrefix = "cite";
const CitationTurnPrefix = "turn";
const CitationSearchPrefix = "search";
const PrivateDelimiterSource = String.raw`\p{Co}\p{Cf}`;
const PrivateCitationMarkerSource = String.raw`[${PrivateDelimiterSource}]*(?:${CitationPrefix}[${PrivateDelimiterSource}]*(?:${CitationTurnPrefix}\d+${CitationSearchPrefix}\d+[${PrivateDelimiterSource}]*)+)[${PrivateDelimiterSource}]*`;
const PrivateCitationMarkerPattern = new RegExp(
  PrivateCitationMarkerSource,
  "giu",
);
const PrivateDelimiterCharacterPattern = new RegExp(
  `[${PrivateDelimiterSource}]`,
  "u",
);
const MaxIncrementalCitationHoldback = 64;

const SearchSemanticTypes = new Set([
  "internalsearchquery",
  "semanticserp",
  "searchquery",
  "referenceslistcomplete",
]);

/**
 * Removes only the private citation markers emitted by the M365 substrate.
 * Private-use and format characters are delimiters in the observed marker
 * form; real URLs and markdown links do not match this pattern.
 */
export function stripPrivateCitationMarkers(text: string): string {
  const matches = [...text.matchAll(PrivateCitationMarkerPattern)];
  if (matches.length === 0) {
    return text;
  }

  let output = "";
  let cursor = 0;
  for (const match of matches) {
    const matchStart = match.index ?? cursor;
    output += text.slice(cursor, matchStart);

    let beforeStart = output.length;
    while (
      beforeStart > 0 &&
      isHorizontalWhitespace(output[beforeStart - 1] ?? "")
    ) {
      beforeStart -= 1;
    }
    const hadBeforeWhitespace = beforeStart < output.length;
    output = output.slice(0, beforeStart);

    let afterEnd = matchStart + match[0].length;
    const afterStart = afterEnd;
    while (isHorizontalWhitespace(text[afterEnd] ?? "")) {
      afterEnd += 1;
    }
    const hadAfterWhitespace = afterEnd > afterStart;
    if (
      hadBeforeWhitespace &&
      hadAfterWhitespace &&
      !isPunctuation(text[afterEnd] ?? "")
    ) {
      output += " ";
    }
    cursor = afterEnd;
  }
  return output + text.slice(cursor);
}

export class IncrementalPrivateCitationMarkerSanitizer {
  private pending = "";

  push(chunk: string): string {
    if (!chunk) {
      return "";
    }
    this.pending += chunk;

    const holdbackStart = findIncrementalHoldbackStart(this.pending);
    const forcedReleaseStart = Math.max(
      0,
      this.pending.length - MaxIncrementalCitationHoldback,
    );
    const releaseEnd = Math.max(holdbackStart, forcedReleaseStart);
    if (releaseEnd <= 0) {
      return "";
    }

    const release = this.pending.slice(0, releaseEnd);
    this.pending = this.pending.slice(releaseEnd);
    const sanitized = stripPrivateCitationMarkers(release);
    if (
      sanitized !== release &&
      /^[ \t]/.test(this.pending) &&
      /[ \t]$/.test(sanitized)
    ) {
      return sanitized.replace(/[ \t]+$/g, "");
    }
    return sanitized;
  }

  finish(): string {
    const remaining = this.pending;
    this.pending = "";
    return stripPrivateCitationMarkers(remaining);
  }
}

/**
 * Extracts only source records found under an observed search semantic
 * message type. The reference object's `referenceId`, `id`, `url`, and
 * `title` shape is unverified because SignalR payload fields are redacted;
 * this narrow decoder is intentionally tolerated and returns no citation for
 * anything else.
 */
export function decodeSubstrateUrlCitations(
  payload: JsonValue | null | undefined,
  assistantText: string,
): JsonObject[] {
  if (payload === null || payload === undefined || !assistantText) {
    return [];
  }

  const ledger = new Map<string, JsonObject>();
  visitForReferences(payload, false, assistantText, ledger);
  return [...ledger.values()];
}

function visitForReferences(
  value: JsonValue,
  inSearchSemanticFrame: boolean,
  assistantText: string,
  ledger: Map<string, JsonObject>,
): void {
  if (Array.isArray(value)) {
    for (const child of value) {
      visitForReferences(child, inSearchSemanticFrame, assistantText, ledger);
    }
    return;
  }
  if (!isJsonObject(value)) {
    return;
  }

  const semanticType =
    tryGetString(value, "messageType") ?? tryGetString(value, "type");
  const searchSemanticFrame =
    inSearchSemanticFrame || isSearchSemanticType(semanticType);

  if (searchSemanticFrame) {
    const citation = decodeReference(value, assistantText);
    if (citation) {
      const identity =
        tryGetString(value, "referenceId") ??
        tryGetString(value, "id") ??
        tryGetString(value, "url");
      if (identity && !ledger.has(identity)) {
        ledger.set(identity, citation);
      }
    }
  }

  for (const child of Object.values(value)) {
    visitForReferences(child, searchSemanticFrame, assistantText, ledger);
  }
}

function isSearchSemanticType(value: string | null): boolean {
  return value !== null && SearchSemanticTypes.has(value.trim().toLowerCase());
}

function decodeReference(
  value: JsonObject,
  assistantText: string,
): JsonObject | null {
  const url = tryGetString(value, "url");
  const title = tryGetString(value, "title");
  if (!url || !title || !isAbsoluteHttpUrl(url) || !title.trim()) {
    return null;
  }
  if (stripPrivateCitationMarkers(url) !== url) {
    return null;
  }

  const citedText = title.trim();
  const titleStart = assistantText.indexOf(citedText);
  const urlStart =
    titleStart >= 0 ? titleStart : assistantText.indexOf(url);
  if (urlStart < 0) {
    return null;
  }
  const spanText = titleStart >= 0 ? citedText : url;
  return {
    type: "url_citation",
    start_index: urlStart,
    end_index: urlStart + spanText.length,
    url,
    title: citedText,
  };
}

function isAbsoluteHttpUrl(value: string): boolean {
  try {
    const parsed = new URL(value);
    return (
      (parsed.protocol === "http:" || parsed.protocol === "https:") &&
      !parsed.username &&
      !parsed.password
    );
  } catch {
    return false;
  }
}

function findIncrementalHoldbackStart(value: string): number {
  if (!value) {
    return 0;
  }

  const suffixLength = findPotentialCitationSuffixLength(value);
  if (suffixLength > 0) {
    let start = value.length - suffixLength;
    while (start > 0 && isHorizontalWhitespace(value[start - 1] ?? "")) {
      start -= 1;
    }
    return start;
  }

  let start = value.length;
  while (start > 0 && isHorizontalWhitespace(value[start - 1] ?? "")) {
    start -= 1;
  }
  return start === value.length ? value.length : start;
}

function findPotentialCitationSuffixLength(value: string): number {
  const searchStart = Math.max(
    0,
    value.length - MaxIncrementalCitationHoldback,
  );
  for (let start = searchStart; start < value.length; start += 1) {
    if (isPotentialCitationMarkerPrefix(value.slice(start))) {
      return value.length - start;
    }
  }
  return 0;
}

function isPotentialCitationMarkerPrefix(value: string): boolean {
  if (!value) {
    return false;
  }
  const normalized = value.toLowerCase();
  let index = 0;
  while (isPrivateDelimiter(normalized[index] ?? "")) {
    index += 1;
  }
  if (index === normalized.length) {
    return true;
  }

  const cite = consumeLiteralPrefix(normalized, index, CitationPrefix);
  if (cite.kind === "mismatch") {
    return false;
  }
  if (cite.kind === "partial") {
    return true;
  }
  index = cite.index;

  while (isPrivateDelimiter(normalized[index] ?? "")) {
    index += 1;
  }
  if (index === normalized.length) {
    return true;
  }

  while (true) {
    const turn = consumeLiteralPrefix(normalized, index, CitationTurnPrefix);
    if (turn.kind === "mismatch") {
      return false;
    }
    if (turn.kind === "partial") {
      return true;
    }
    index = turn.index;

    const turnDigitsStart = index;
    while (isAsciiDigit(normalized[index] ?? "")) {
      index += 1;
    }
    if (index === turnDigitsStart || index === normalized.length) {
      return index === normalized.length;
    }

    const search = consumeLiteralPrefix(
      normalized,
      index,
      CitationSearchPrefix,
    );
    if (search.kind === "mismatch") {
      return false;
    }
    if (search.kind === "partial") {
      return true;
    }
    index = search.index;

    const searchDigitsStart = index;
    while (isAsciiDigit(normalized[index] ?? "")) {
      index += 1;
    }
    if (index === searchDigitsStart || index === normalized.length) {
      return index === normalized.length;
    }

    while (isPrivateDelimiter(normalized[index] ?? "")) {
      index += 1;
    }
    if (index === normalized.length) {
      return true;
    }
  }
}

function consumeLiteralPrefix(
  value: string,
  start: number,
  literal: string,
): { kind: "complete"; index: number } | { kind: "partial" } | { kind: "mismatch" } {
  for (let offset = 0; offset < literal.length; offset += 1) {
    const index = start + offset;
    if (index >= value.length) {
      return { kind: "partial" };
    }
    if (value[index] !== literal[offset]) {
      return { kind: "mismatch" };
    }
  }
  return { kind: "complete", index: start + literal.length };
}

function isPrivateDelimiter(value: string): boolean {
  return value.length === 1 && PrivateDelimiterCharacterPattern.test(value);
}

function isAsciiDigit(value: string): boolean {
  return value >= "0" && value <= "9";
}

function isHorizontalWhitespace(value: string): boolean {
  return value === " " || value === "\t";
}

function isPunctuation(value: string): boolean {
  return ",.;:!?".includes(value);
}
