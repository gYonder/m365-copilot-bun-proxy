import { createHash, randomUUID } from "node:crypto";
import type { JsonObject, JsonValue } from "./types";

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
  | "cancellation";

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
  private sequence = 0;

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
      recentEvents: this.recentEvents.map((event) => ({
        sequence: event.sequence,
        name: event.name,
        correlationId: event.correlationId,
        occurredAtUnix: event.occurredAtUnix,
        fields: { ...event.fields },
      })),
    };
  }
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
