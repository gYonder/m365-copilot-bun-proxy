import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import packageMetadata from "../../package.json";

const StateSchemaVersion = 1;
const BridgeContractVersion = 2;
const AdapterRoute = "substrate-coding";
const ProviderRuntimeVersion = packageMetadata.version;
const MaxDurableEntries = 1_024;

type DurableResponseEntry = {
  conversationId: string;
  expiresAtUtc: number;
  contextInputTokens?: number;
  contextWindowId?: string | null;
};

type DurableConversationEntry = {
  conversationId: string;
  expiresAtUtc: number;
};

type DurableSessionEntry = {
  sessionId: string;
  expiresAtUtc: number;
};

export type DurableReplayEntry = {
  conversationId: string | null;
  expiresAtUtc: number;
  responseFingerprint?: string;
};

export type DurableToolLedgerEntry = {
  serialized: string;
  expiresAtUtc: number;
};

export type DurableState = {
  schema_version: typeof StateSchemaVersion;
  contract_version: typeof BridgeContractVersion;
  adapter_route: typeof AdapterRoute;
  provider_runtime_version: string;
  responses: Record<string, DurableResponseEntry>;
  conversations: Record<string, DurableConversationEntry>;
  sessions: Record<string, DurableSessionEntry>;
  replays: Record<string, DurableReplayEntry>;
  toolLedgers: Record<string, DurableToolLedgerEntry>;
};

type QuarantineStatus = {
  count: number;
  lastOccurredAtUnix: number | null;
};

class DurableStateParseError extends Error {
  constructor(
    readonly reason: string,
    readonly found: Record<string, unknown> = {},
  ) {
    super(reason);
  }
}

const emptyState = (): DurableState => ({
  schema_version: StateSchemaVersion,
  contract_version: BridgeContractVersion,
  adapter_route: AdapterRoute,
  provider_runtime_version: ProviderRuntimeVersion,
  responses: {},
  conversations: {},
  sessions: {},
  replays: {},
  toolLedgers: {},
});

export function durableStatePath(): string | null {
  const configured = process.env.M365_RUNTIME_STATE_PATH?.trim();
  return configured || null;
}

export class DurableStateStore {
  readonly path: string | null;
  state: DurableState;
  private quarantineCountValue = 0;
  private lastQuarantineAtUnixValue: number | null = null;

  constructor(filePath = durableStatePath()) {
    this.path = filePath;
    this.state = this.load();
  }

  get quarantineStatus(): QuarantineStatus {
    return {
      count: this.quarantineCountValue,
      lastOccurredAtUnix: this.lastQuarantineAtUnixValue,
    };
  }

  private load(): DurableState {
    if (!this.path || !existsSync(this.path)) return emptyState();
    try {
      return parseState(JSON.parse(readFileSync(this.path, "utf8")));
    } catch (error) {
      if (!existsSync(this.path)) {
        return emptyState();
      }
      const failure =
        error instanceof DurableStateParseError
          ? error
          : new DurableStateParseError("unparseable_state");
      this.quarantine(failure);
      return emptyState();
    }
  }

  save(): void {
    if (!this.path) return;
    const now = Date.now();
    for (const collection of [
      this.state.responses,
      this.state.conversations,
      this.state.sessions,
      this.state.replays,
      this.state.toolLedgers,
    ]) {
      for (const [key, value] of Object.entries(collection)) {
        if (value.expiresAtUtc <= now) delete collection[key];
      }
      const keys = Object.keys(collection);
      for (const key of keys.slice(
        0,
        Math.max(0, keys.length - MaxDurableEntries),
      )) {
        delete collection[key];
      }
    }
    const dir = path.dirname(this.path);
    mkdirSync(dir, { recursive: true, mode: 0o700 });
    const temp = `${this.path}.${process.pid}.tmp`;
    writeFileSync(temp, JSON.stringify(this.state) + "\n", { mode: 0o600 });
    chmodSync(temp, 0o600);
    renameSync(temp, this.path);
    chmodSync(this.path, 0o600);
  }

  private quarantine(error: DurableStateParseError): void {
    this.quarantineCountValue += 1;
    this.lastQuarantineAtUnixValue = Math.floor(Date.now() / 1000);

    let quarantinePath: string | null = null;
    if (this.path) {
      const base = `${this.path}.incompatible-${Date.now()}`;
      let candidate = base;
      let suffix = 0;
      while (existsSync(candidate)) {
        suffix += 1;
        candidate = `${base}-${suffix}`;
      }
      try {
        renameSync(this.path, candidate);
        quarantinePath = candidate;
      } catch {
        quarantinePath = null;
      }
    }

    console.error(
      JSON.stringify({
        event: "durable_state_quarantined",
        path: this.path,
        quarantine_path: quarantinePath,
        reason: error.reason,
        expected: {
          schema_version: StateSchemaVersion,
          contract_version: BridgeContractVersion,
          adapter_route: AdapterRoute,
        },
        found: {
          schema_version: boundedVersionForLog(error.found.schema_version),
          contract_version: boundedVersionForLog(error.found.contract_version),
          adapter_route: boundedVersionForLog(error.found.adapter_route),
        },
      }),
    );
  }
}

function parseState(value: unknown): DurableState {
  if (!isRecord(value)) {
    throw new DurableStateParseError("state_root_not_object");
  }

  const found = {
    schema_version: value.schema_version ?? null,
    contract_version: value.contract_version ?? null,
    adapter_route: value.adapter_route ?? null,
  };
  if (value.schema_version !== StateSchemaVersion) {
    throw new DurableStateParseError("unsupported_schema_version", found);
  }
  if (value.contract_version !== BridgeContractVersion) {
    throw new DurableStateParseError("unsupported_contract_version", found);
  }
  if (value.adapter_route !== AdapterRoute) {
    throw new DurableStateParseError("unsupported_adapter_route", found);
  }
  if (
    typeof value.provider_runtime_version !== "string" ||
    !value.provider_runtime_version.trim()
  ) {
    throw new DurableStateParseError("missing_provider_runtime_version", found);
  }

  return {
    schema_version: StateSchemaVersion,
    contract_version: BridgeContractVersion,
    adapter_route: AdapterRoute,
    provider_runtime_version: value.provider_runtime_version.trim(),
    responses: parseResponseEntries(value.responses),
    conversations: parseConversationEntries(value.conversations),
    sessions: parseSessionEntries(value.sessions),
    replays: parseReplayEntries(value.replays),
    toolLedgers: parseToolLedgerEntries(value.toolLedgers),
  };
}

function parseToolLedgerEntries(
  value: unknown,
): Record<string, DurableToolLedgerEntry> {
  const record = requireRecord(value, "tool_ledger_collection");
  const output: Record<string, DurableToolLedgerEntry> = {};
  for (const [key, entry] of Object.entries(record)) {
    const item = requireRecord(entry, "tool_ledger_entry");
    const serialized = requireString(item.serialized, "tool_ledger_serialized");
    const expiresAtUtc = requirePositiveNumber(
      item.expiresAtUtc,
      "tool_ledger_expiry",
    );
    output[key] = { serialized, expiresAtUtc };
  }
  return output;
}

function parseReplayEntries(value: unknown): Record<string, DurableReplayEntry> {
  const record = requireRecord(value, "replay_collection");
  const output: Record<string, DurableReplayEntry> = {};
  for (const [key, entry] of Object.entries(record)) {
    const item = requireRecord(entry, "replay_entry");
    if (Object.hasOwn(item, "response")) {
      throw new DurableStateParseError("replay_response_body_present");
    }
    const expiresAtUtc = requirePositiveNumber(
      item.expiresAtUtc,
      "replay_expiry",
    );
    let conversationId: string | null;
    if (item.conversationId === null) {
      conversationId = null;
    } else {
      conversationId = requireString(item.conversationId, "replay_conversation");
    }

    const responseFingerprint =
      item.responseFingerprint === undefined
        ? undefined
        : requireBoundedString(
            item.responseFingerprint,
            "replay_fingerprint",
            128,
          );
    output[key] = {
      conversationId,
      expiresAtUtc,
      ...(responseFingerprint === undefined ? {} : { responseFingerprint }),
    };
  }
  return output;
}

function parseResponseEntries(value: unknown): Record<string, DurableResponseEntry> {
  const record = requireRecord(value, "response_collection");
  const output: Record<string, DurableResponseEntry> = {};
  for (const [key, entry] of Object.entries(record)) {
    const item = requireRecord(entry, "response_entry");
    const conversationId = requireString(
      item.conversationId,
      "response_conversation",
    );
    const expiresAtUtc = requirePositiveNumber(
      item.expiresAtUtc,
      "response_expiry",
    );
    const contextInputTokens =
      item.contextInputTokens === undefined
        ? undefined
        : requirePositiveNumber(
            item.contextInputTokens,
            "response_context_tokens",
          );
    const contextWindowId =
      item.contextWindowId === undefined
        ? undefined
        : item.contextWindowId === null
          ? null
          : requireString(item.contextWindowId, "response_context_window");
    output[key] = {
      conversationId,
      expiresAtUtc,
      ...(contextInputTokens === undefined ? {} : { contextInputTokens }),
      ...(contextWindowId === undefined ? {} : { contextWindowId }),
    };
  }
  return output;
}

function parseConversationEntries(
  value: unknown,
): Record<string, DurableConversationEntry> {
  const record = requireRecord(value, "conversation_collection");
  const output: Record<string, DurableConversationEntry> = {};
  for (const [key, entry] of Object.entries(record)) {
    const item = requireRecord(entry, "conversation_entry");
    output[key] = {
      conversationId: requireString(
        item.conversationId,
        "conversation_id",
      ),
      expiresAtUtc: requirePositiveNumber(
        item.expiresAtUtc,
        "conversation_expiry",
      ),
    };
  }
  return output;
}

function parseSessionEntries(value: unknown): Record<string, DurableSessionEntry> {
  const record = requireRecord(value, "session_collection");
  const output: Record<string, DurableSessionEntry> = {};
  for (const [key, entry] of Object.entries(record)) {
    const item = requireRecord(entry, "session_entry");
    output[key] = {
      sessionId: requireString(item.sessionId, "session_id"),
      expiresAtUtc: requirePositiveNumber(item.expiresAtUtc, "session_expiry"),
    };
  }
  return output;
}

function requireRecord(value: unknown, reason: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new DurableStateParseError(reason);
  }
  return value;
}

function requireString(value: unknown, reason: string): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new DurableStateParseError(reason);
  }
  return value.trim();
}

function requireBoundedString(
  value: unknown,
  reason: string,
  maxLength: number,
): string {
  const normalized = requireString(value, reason);
  if (normalized.length > maxLength) {
    throw new DurableStateParseError(reason);
  }
  return normalized;
}

function requirePositiveNumber(value: unknown, reason: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new DurableStateParseError(reason);
  }
  return value;
}

function boundedVersionForLog(value: unknown): string | number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") return value.slice(0, 64);
  return null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
