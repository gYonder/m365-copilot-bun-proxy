import {
  ProxyVizTraceStatuses,
  type JsonValue,
  type ProxyVizTraceRecord,
} from "./types";
import { cloneJsonValue, nowUnix } from "./utils";

type MutableTraceRecord = ProxyVizTraceRecord & {
  expiresAtUnix: number;
};

const MaxTraceEntries = 256;

export class ProxyVizTraceStore {
  private readonly traces = new Map<string, MutableTraceRecord>();
  private readonly aliases = new Map<string, string>();

  constructor(private readonly ttlSeconds: number) {}

  start(
    traceId: string,
    requestType: string,
    transformMode: string,
    transport: string,
  ): void {
    const now = nowUnix();
    this.traces.set(traceId, {
      traceId,
      status: ProxyVizTraceStatuses.Pending,
      requestType,
      transformMode,
      transport,
      proxyStatusCode: null,
      upstreamStatusCode: null,
      pane2: null,
      pane3: null,
      pane4: null,
      error: null,
      createdAtUnix: now,
      updatedAtUnix: now,
      expiresAtUnix: now + this.ttlSeconds,
    });
    this.cleanup(now);
    while (this.traces.size > MaxTraceEntries) {
      const oldest = this.traces.keys().next();
      if (oldest.done) {
        break;
      }

      this.traces.delete(oldest.value);
    }
  }

  alias(requestedTraceId: string, serverTraceId: string): void {
    const alias = requestedTraceId.trim();
    const traceId = serverTraceId.trim();
    if (alias && traceId) {
      this.aliases.set(alias, traceId);
    }
  }

  setPane3(traceId: string, pane3: JsonValue | null): void {
    this.update(traceId, (record) => {
      record.pane3 = cloneOrNull(pane3);
    });
  }

  setPane4(
    traceId: string,
    pane4: JsonValue | null,
    upstreamStatusCode: number | null = null,
  ): void {
    this.update(traceId, (record) => {
      record.pane4 = cloneOrNull(pane4);
      if (upstreamStatusCode !== null) {
        record.upstreamStatusCode = upstreamStatusCode;
      }
    });
  }

  setPane2(
    traceId: string,
    pane2: JsonValue | null,
    proxyStatusCode: number | null = null,
  ): void {
    this.update(traceId, (record) => {
      record.pane2 = cloneOrNull(pane2);
      if (proxyStatusCode !== null) {
        record.proxyStatusCode = proxyStatusCode;
      }
    });
  }

  setError(
    traceId: string,
    error: JsonValue | null,
    proxyStatusCode: number | null = null,
  ): void {
    this.update(traceId, (record) => {
      record.error = cloneOrNull(error);
      record.status = ProxyVizTraceStatuses.Failed;
      if (proxyStatusCode !== null) {
        record.proxyStatusCode = proxyStatusCode;
      }
    });
  }

  setUpstreamStatus(traceId: string, upstreamStatusCode: number): void {
    this.update(traceId, (record) => {
      record.upstreamStatusCode = upstreamStatusCode;
    });
  }

  complete(traceId: string, proxyStatusCode: number | null = null): void {
    this.update(traceId, (record) => {
      record.status = ProxyVizTraceStatuses.Completed;
      if (proxyStatusCode !== null) {
        record.proxyStatusCode = proxyStatusCode;
      }
    });
  }

  fail(traceId: string, proxyStatusCode: number | null = null): void {
    this.update(traceId, (record) => {
      record.status = ProxyVizTraceStatuses.Failed;
      if (proxyStatusCode !== null) {
        record.proxyStatusCode = proxyStatusCode;
      }
    });
  }

  get(traceId: string): ProxyVizTraceRecord | null {
    const now = nowUnix();
    this.cleanup(now);
    const existing = this.traces.get(traceId) ??
      this.traces.get(this.aliases.get(traceId) ?? "");
    if (!existing) {
      return null;
    }
    const { expiresAtUnix: _expiresAtUnix, ...record } = existing;
    return cloneJsonValue(record);
  }

  private update(
    traceId: string,
    updater: (record: MutableTraceRecord) => void,
  ): void {
    const existing = this.traces.get(traceId);
    if (!existing) {
      return;
    }
    updater(existing);
    existing.updatedAtUnix = nowUnix();
    existing.expiresAtUnix = existing.updatedAtUnix + this.ttlSeconds;
  }

  private cleanup(now: number): void {
    for (const [traceId, record] of this.traces.entries()) {
      if (record.expiresAtUnix < now) {
        this.traces.delete(traceId);
      }
    }
    for (const [alias, traceId] of this.aliases.entries()) {
      if (!this.traces.has(traceId)) {
        this.aliases.delete(alias);
      }
    }
  }
}

function cloneOrNull(value: JsonValue | null): JsonValue | null {
  if (value === null) {
    return null;
  }
  return cloneJsonValue(value);
}
