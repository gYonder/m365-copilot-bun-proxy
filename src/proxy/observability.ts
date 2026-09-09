import { createHash, randomUUID } from "node:crypto";
import {
  appendFileSync,
  chmodSync,
  existsSync,
  mkdirSync,
  readdirSync,
  renameSync,
  statSync,
  unlinkSync,
} from "node:fs";
import path from "node:path";
import type {
  JsonObject,
  JsonValue,
  ObservabilityOptions,
} from "./types";

export type BridgeEventName =
  | "auth_path"
  | "queue_wait"
  | "attempt_outcome"
  | "retry"
  | "substrate_terminal"
  | "dedup_hit"
  | "store_eviction"
  | "truncation"
  | "image_quota"
  | "cancellation"
  | "tool_ledger_recovery"
  | "provider_drift"
  | "attempted_tool_call_in_unguarded_stream"
  | "tool_call_recovery_exhausted";

export type BridgeEvent = {
  sequence: number;
  name: BridgeEventName;
  correlationId: string;
  occurredAtUnix: number;
  fields: JsonObject;
};

type EventSummary = {
  count: number;
  lastOccurredAtUnix: number;
};

const MaxRecentEvents = 64;
const ForbiddenKeyPattern =
  /(?:authorization|token|cookie|prompt|body|account|object.?id|tenant.?id|browser|websocket|signed.?url|artifact.?url|access.?token|refresh.?token)/i;
const SensitiveValuePattern =
  /(?:bearer\s+[a-z0-9._~-]+|(?:access_token|sig|signature|token)=[^&\s]+)/gi;

export class BridgeObservability {
  private readonly startedAtUnix = Math.floor(Date.now() / 1000);
  private readonly recentEvents: BridgeEvent[] = [];
  private readonly summaries = new Map<BridgeEventName, EventSummary>();
  private readonly eventLog: BridgeEventLog | null;
  private eventLogHealthy = true;
  private eventLogWriteFailures = 0;
  private sequence = 0;

  constructor(options?: ObservabilityOptions) {
    this.eventLog = options?.enabled ? new BridgeEventLog(options) : null;
  }

  createCorrelationId(seed?: string | null): string {
    const normalized = seed?.trim();
    if (!normalized) return `corr_${randomUUID().replaceAll("-", "")}`;
    return `corr_${createHash("sha256").update(normalized).digest("hex").slice(0, 24)}`;
  }

  record(
    name: BridgeEventName,
    fields: JsonObject = {},
    correlationId?: string | null,
  ): BridgeEvent {
    const occurredAtUnix = Math.floor(Date.now() / 1000);
    const event: BridgeEvent = {
      sequence: ++this.sequence,
      name,
      correlationId: this.createCorrelationId(correlationId),
      occurredAtUnix,
      fields: sanitizeTelemetryObject(fields),
    };
    this.recentEvents.push(event);
    while (this.recentEvents.length > MaxRecentEvents) this.recentEvents.shift();
    const previous = this.summaries.get(name);
    this.summaries.set(name, {
      count: (previous?.count ?? 0) + 1,
      lastOccurredAtUnix: occurredAtUnix,
    });
    this.persist(event);
    return event;
  }

  readiness(): JsonObject {
    const events: JsonObject = {};
    for (const [name, summary] of this.summaries) {
      events[name] = { ...summary };
    }
    return {
      status: "ready",
      startedAtUnix: this.startedAtUnix,
      eventSequence: this.sequence,
      events,
      eventLog: {
        enabled: this.eventLog !== null,
        healthy: this.eventLogHealthy,
        writeFailures: this.eventLogWriteFailures,
      },
      recentEvents: this.recentEvents.map((event) => ({
        sequence: event.sequence,
        name: event.name,
        correlationId: event.correlationId,
        occurredAtUnix: event.occurredAtUnix,
        fields: { ...event.fields },
      })),
    };
  }

  private persist(event: BridgeEvent): void {
    if (!this.eventLog) {
      return;
    }
    try {
      this.eventLog.write(event);
      this.eventLogHealthy = true;
    } catch (error) {
      this.eventLogHealthy = false;
      this.eventLogWriteFailures += 1;
      console.error(
        `[observability] failed to persist bridge event (${safeErrorCode(error)})`,
      );
    }
  }
}

class BridgeEventLog {
  constructor(private readonly options: ObservabilityOptions) {}

  write(event: BridgeEvent): void {
    const line = this.serialize(event);
    mkdirSync(path.dirname(this.options.logPath), {
      recursive: true,
      mode: 0o700,
    });
    if (this.shouldRotate(line)) {
      this.rotate();
    }
    appendFileSync(this.options.logPath, line, {
      encoding: "utf8",
      mode: 0o600,
    });
    chmodSync(this.options.logPath, 0o600);
  }

  private serialize(event: BridgeEvent): string {
    const line = `${JSON.stringify(event)}\n`;
    const originalBytes = Buffer.byteLength(line);
    if (originalBytes <= this.options.maxBytes) {
      return line;
    }
    const compactLine = `${JSON.stringify({
      ...event,
      fields: { truncated: true, originalBytes },
    })}\n`;
    if (Buffer.byteLength(compactLine) > this.options.maxBytes) {
      throw new Error("Configured observability maxBytes is too small.");
    }
    return compactLine;
  }

  private shouldRotate(line: string): boolean {
    if (!existsSync(this.options.logPath)) {
      return false;
    }
    return (
      statSync(this.options.logPath).size + Buffer.byteLength(line) >
      this.options.maxBytes
    );
  }

  private rotate(): void {
    this.removeStaleArchives();
    if (this.options.maxFiles === 1) {
      unlinkSync(this.options.logPath);
      return;
    }
    for (let index = this.options.maxFiles - 1; index >= 1; index -= 1) {
      const target = `${this.options.logPath}.${index}`;
      if (existsSync(target)) {
        unlinkSync(target);
      }
      const source =
        index === 1
          ? this.options.logPath
          : `${this.options.logPath}.${index - 1}`;
      if (existsSync(source)) {
        renameSync(source, target);
      }
    }
  }

  private removeStaleArchives(): void {
    const directory = path.dirname(this.options.logPath);
    const baseName = path.basename(this.options.logPath);
    for (const entry of readdirSync(directory)) {
      const match = new RegExp(`^${escapeRegExp(baseName)}\\.(\\d+)$`).exec(
        entry,
      );
      if (match && Number(match[1]) >= this.options.maxFiles) {
        unlinkSync(path.join(directory, entry));
      }
    }
  }
}

function safeErrorCode(error: unknown): string {
  if (
    error &&
    typeof error === "object" &&
    "code" in error &&
    typeof error.code === "string"
  ) {
    return error.code.replace(/[^A-Za-z0-9_-]/g, "").slice(0, 64) || "unknown";
  }
  return "unknown";
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function sanitizeTelemetryObject(value: JsonObject): JsonObject {
  const sanitized: JsonObject = {};
  for (const [key, child] of Object.entries(value)) {
    sanitized[key] = sanitizeTelemetryValue(key, child);
  }
  return sanitized;
}

function sanitizeTelemetryValue(key: string, value: JsonValue): JsonValue {
  if (ForbiddenKeyPattern.test(key)) return "[redacted]";
  if (typeof value === "string") {
    if (/^(?:wss?|https?):\/\//i.test(value)) return "[redacted-url]";
    return value.replace(SensitiveValuePattern, "[redacted]");
  }
  if (Array.isArray(value)) {
    return value.map((item) => sanitizeTelemetryValue(key, item));
  }
  if (value && typeof value === "object") {
    return sanitizeTelemetryObject(value as JsonObject);
  }
  return value;
}
