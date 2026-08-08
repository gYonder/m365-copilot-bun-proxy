import { tryCanonicalizeJson } from "./canonical-json";
import {
  classifyBridgeFailure,
  type BridgeFailure,
  type InternalReason,
} from "./failure-classifier";

// Keep this at or below response-store.ts:31's 60-second replay guard.
export const TOOL_LEDGER_PENDING_TTL_MS = 60_000;
export const MAX_TOOL_CALLS_PER_RESPONSE = 8;
export const MAX_TOOL_ROUNDS_PER_TASK = 8;
export const MAX_TOOL_LEDGER_ENTRIES = 128;
export const MAX_TOOL_REPETITIONS_PER_TASK = 2;
export const MAX_TOOL_CORRECTIVE_REASKS_PER_CALL = 1;
export const TOOL_LEDGER_STATE_VERSION = 1;

export type IssuedCall = {
  call_id: string;
  name: string;
  type: string;
  canonical_arguments: string;
  request_profile_key: string;
  issued_at: number;
  status: "pending" | "completed";
};

export type ToolCallToIssue = {
  call_id: string;
  name: string;
  type: string;
  arguments: unknown;
};

export type IssueCallsRequest = {
  taskId: string;
  responseId: string;
  requestProfileKey: string;
  calls: readonly ToolCallToIssue[];
  round: number;
};

export type ToolLedgerOptions = {
  now?: () => number;
  pendingTtlMs?: number;
  maxCallsPerResponse?: number;
  maxRoundsPerTask?: number;
  maxEntries?: number;
  maxRepetitionsPerTask?: number;
};

export type ToolLedgerRejectionKind =
  | "invalid_call"
  | "call_cap_exceeded"
  | "round_cap_exceeded"
  | "repetition_bound_exhausted"
  | "ledger_capacity_exceeded"
  | "unknown_call_id"
  | "duplicate_result"
  | "cross_profile"
  | "out_of_order"
  | "corrective_reask_exhausted";

export class ToolLedgerConfigurationError extends Error {
  readonly name = "ToolLedgerConfigurationError";
}

export class ToolLedgerCodecError extends Error {
  readonly name = "ToolLedgerCodecError";

  constructor(
    readonly code:
      | "invalid_serialization"
      | "unsupported_version"
      | "malformed_state",
    message: string,
  ) {
    super(message);
  }
}

export class ToolLedgerError extends Error {
  readonly name = "ToolLedgerError";
  readonly code: BridgeFailure["code"];
  readonly reason: InternalReason;

  constructor(
    readonly kind: ToolLedgerRejectionKind,
    readonly failure: BridgeFailure,
  ) {
    super(failure.message);
    this.code = failure.code;
    this.reason = failure.reason;
  }
}

export type ToolLedgerResult<T> =
  | { ok: true; value: T }
  | { ok: false; error: ToolLedgerError };

export type ToolLedgerRecoveryOutcome = {
  taskId: string;
  callIds: string[];
  failure: BridgeFailure;
  released: true;
};

export type ToolParallelismMetadata = {
  requestPermitsParallelCalls: boolean;
  independentlyReadOnly: boolean;
  consumesAnotherCallOutput: boolean;
};

export type ToolParallelismDecision = {
  parallel: boolean;
  reason:
    | "allowed"
    | "independence_signal_unavailable"
    | "request_disallows_parallel"
    | "mutating_or_dependency_sensitive"
    | "call_bound_exceeded";
};

type CallMetadata = {
  taskId: string;
  responseId: string;
  round: number;
  sequence: number;
};

type TaskState = {
  rounds: Set<number>;
  repetitions: Map<string, number>;
  responseCounts: Map<string, number>;
  correctiveReaskCallIds: Set<string>;
};

type PersistedState = {
  version: number;
  entries: IssuedCall[];
  metadata: Array<CallMetadata & { call_id: string }>;
  tasks: Array<{
    task_id: string;
    rounds: number[];
    repetitions: Array<{
      name: string;
      canonical_arguments: string;
      count: number;
    }>;
    response_counts: Array<{ key: string; count: number }>;
    corrective_reask_call_ids: string[];
  }>;
};

export class ToolLedger {
  private readonly entries = new Map<string, IssuedCall>();
  private readonly metadata = new Map<string, CallMetadata>();
  private readonly tasks = new Map<string, TaskState>();
  private readonly now: () => number;
  private readonly pendingTtlMs: number;
  private readonly maxCallsPerResponse: number;
  private readonly maxRoundsPerTask: number;
  private readonly maxEntries: number;
  private readonly maxRepetitionsPerTask: number;
  private sequence = 0;

  constructor(options: ToolLedgerOptions = {}) {
    this.now = options.now ?? Date.now;
    this.pendingTtlMs = options.pendingTtlMs ?? TOOL_LEDGER_PENDING_TTL_MS;
    this.maxCallsPerResponse =
      options.maxCallsPerResponse ?? MAX_TOOL_CALLS_PER_RESPONSE;
    this.maxRoundsPerTask =
      options.maxRoundsPerTask ?? MAX_TOOL_ROUNDS_PER_TASK;
    this.maxEntries = options.maxEntries ?? MAX_TOOL_LEDGER_ENTRIES;
    this.maxRepetitionsPerTask =
      options.maxRepetitionsPerTask ?? MAX_TOOL_REPETITIONS_PER_TASK;
    bound("pendingTtlMs", this.pendingTtlMs, 1, TOOL_LEDGER_PENDING_TTL_MS);
    bound("maxCallsPerResponse", this.maxCallsPerResponse);
    bound("maxRoundsPerTask", this.maxRoundsPerTask);
    bound("maxEntries", this.maxEntries);
    bound(
      "maxRepetitionsPerTask",
      this.maxRepetitionsPerTask,
      1,
      MAX_TOOL_REPETITIONS_PER_TASK,
    );
  }

  get size(): number {
    return this.entries.size;
  }

  issueCalls(request: IssueCallsRequest): ToolLedgerResult<IssuedCall[]> {
    const now = this.readNow();
    this.expirePending(now);
    const requestError = validateIssueRequest(request);
    if (requestError) return { ok: false, error: requestError };
    if (request.calls.length > this.maxCallsPerResponse) {
      return this.reject("call_cap_exceeded", "invalid_request");
    }

    const task = this.tasks.get(request.taskId);
    const latestRound = task ? Math.max(...task.rounds, 0) : 0;
    const sameRound = task?.rounds.has(request.round) ?? false;
    if (
      request.round > this.maxRoundsPerTask ||
      (!sameRound && task && request.round <= latestRound) ||
      (!sameRound && (task?.rounds.size ?? 0) >= this.maxRoundsPerTask)
    ) {
      return this.reject("round_cap_exceeded", "invalid_request");
    }

    const responseCount = task?.responseCounts.get(request.responseId) ?? 0;
    if (responseCount + request.calls.length > this.maxCallsPerResponse) {
      return this.reject("call_cap_exceeded", "invalid_request");
    }

    const prepared: Array<{
      input: ToolCallToIssue;
      canonicalArguments: string;
      repetitionKey: string;
    }> = [];
    const callIds = new Set<string>();
    const batchRepetitions = new Map<string, number>();
    for (const input of request.calls) {
      if (
        !isRecord(input) ||
        !nonEmpty(input.call_id) ||
        !nonEmpty(input.name) ||
        !nonEmpty(input.type) ||
        callIds.has(input.call_id) ||
        this.entries.has(input.call_id)
      ) {
        return this.reject("invalid_call", "invalid_request");
      }

      const canonical = canonicalArguments(input.arguments);
      if (!canonical.ok) return canonical;
      const repetitionKey = repetitionIdentity(input.name, canonical.value);
      const previous = task?.repetitions.get(repetitionKey) ?? 0;
      const inBatch = batchRepetitions.get(repetitionKey) ?? 0;
      if (
        previous + inBatch >= this.maxRepetitionsPerTask ||
        (previous === 0 &&
          !batchRepetitions.has(repetitionKey) &&
          (task?.repetitions.size ?? 0) + batchRepetitions.size >=
            MAX_TOOL_LEDGER_ENTRIES)
      ) {
        return this.reject(
          "repetition_bound_exhausted",
          "tool_round_repetition_bound_exhausted",
        );
      }
      callIds.add(input.call_id);
      batchRepetitions.set(repetitionKey, inBatch + 1);
      prepared.push({
        input,
        canonicalArguments: canonical.value,
        repetitionKey,
      });
    }

    if (!this.makeRoom(prepared.length)) {
      return this.reject("ledger_capacity_exceeded", "invalid_request");
    }
    const state = task ?? emptyTask();
    this.tasks.set(request.taskId, state);
    state.rounds.add(request.round);
    state.responseCounts.set(
      request.responseId,
      responseCount + prepared.length,
    );

    const issued: IssuedCall[] = [];
    for (const item of prepared) {
      const entry: IssuedCall = {
        call_id: item.input.call_id,
        name: item.input.name,
        type: item.input.type,
        canonical_arguments: item.canonicalArguments,
        request_profile_key: request.requestProfileKey,
        issued_at: now,
        status: "pending",
      };
      this.entries.set(entry.call_id, entry);
      this.metadata.set(entry.call_id, {
        taskId: request.taskId,
        responseId: request.responseId,
        round: request.round,
        sequence: this.sequence++,
      });
      state.repetitions.set(
        item.repetitionKey,
        (state.repetitions.get(item.repetitionKey) ?? 0) + 1,
      );
      issued.push({ ...entry });
    }
    return { ok: true, value: issued };
  }

  acceptResult(
    callId: string,
    requestProfileKey: string,
    _result?: unknown,
  ): ToolLedgerResult<IssuedCall> {
    this.expirePending(this.readNow());
    const entry = this.entries.get(callId);
    if (!entry) return this.reject("unknown_call_id", "invalid_request");
    if (entry.status === "completed") {
      return this.reject("duplicate_result", "duplicate_tool_result_or_replay");
    }
    if (entry.request_profile_key !== requestProfileKey) {
      return this.reject("cross_profile", "invalid_request");
    }

    const metadata = this.metadata.get(callId);
    const task = metadata ? this.tasks.get(metadata.taskId) : undefined;
    if (!metadata || !task) return this.reject("invalid_call", "invalid_request");

    // Sibling calls in one round are parallel-safe. A pending call from an
    // earlier round is stale once a later round has been issued.
    const latestRound = Math.max(...task.rounds, 0);
    if (metadata.round < latestRound) {
      return this.reject("out_of_order", "invalid_request");
    }
    entry.status = "completed";
    return { ok: true, value: { ...entry } };
  }

  requestCorrectiveReask(
    callId: string,
    requestProfileKey: string,
  ): ToolLedgerResult<{ call_id: string; allowed: true }> {
    this.expirePending(this.readNow());
    const entry = this.entries.get(callId);
    if (!entry) return this.reject("unknown_call_id", "invalid_request");
    if (entry.status === "completed") {
      return this.reject("duplicate_result", "duplicate_tool_result_or_replay");
    }
    if (entry.request_profile_key !== requestProfileKey) {
      return this.reject("cross_profile", "invalid_request");
    }
    const task = this.tasks.get(this.metadata.get(callId)?.taskId ?? "");
    if (!task) return this.reject("invalid_call", "invalid_request");
    const correctiveReasks = task.correctiveReaskCallIds.has(callId) ? 1 : 0;
    if (correctiveReasks >= MAX_TOOL_CORRECTIVE_REASKS_PER_CALL) {
      return this.reject(
        "corrective_reask_exhausted",
        "tool_round_repetition_bound_exhausted",
      );
    }
    task.correctiveReaskCallIds.add(callId);
    return { ok: true, value: { call_id: callId, allowed: true } };
  }

  get(callId: string): IssuedCall | null {
    const entry = this.entries.get(callId);
    return entry ? { ...entry } : null;
  }

  pendingCalls(taskId?: string): IssuedCall[] {
    this.expirePending(this.readNow());
    return [...this.entries.values()]
      .filter(
        (entry) =>
          entry.status === "pending" &&
          (!taskId || this.metadata.get(entry.call_id)?.taskId === taskId),
      )
      .sort(
        (left, right) =>
          this.metadata.get(left.call_id)!.sequence -
          this.metadata.get(right.call_id)!.sequence,
      )
      .map((entry) => ({ ...entry }));
  }

  hasPending(taskId?: string): boolean {
    return this.pendingCalls(taskId).length > 0;
  }

  recoverExpired(now = this.readNow()): ToolLedgerRecoveryOutcome[] {
    return this.expirePending(now);
  }

  serialize(): string {
    const state: PersistedState = {
      version: TOOL_LEDGER_STATE_VERSION,
      entries: [...this.entries.values()].map((entry) => ({ ...entry })),
      metadata: [...this.metadata.entries()]
        .sort(([, left], [, right]) => left.sequence - right.sequence)
        .map(([call_id, metadata]) => ({ call_id, ...metadata })),
      tasks: [...this.tasks.entries()].map(([task_id, task]) => ({
        task_id,
        rounds: [...task.rounds].sort((left, right) => left - right),
        repetitions: [...task.repetitions.entries()].map(
          ([identity, count]) => {
            const [name, canonical_arguments] = JSON.parse(identity) as [
              string,
              string,
            ];
            return { name, canonical_arguments, count };
          },
        ),
        response_counts: [...task.responseCounts.entries()].map(
          ([key, count]) => ({ key, count }),
        ),
        corrective_reask_call_ids: [...task.correctiveReaskCallIds],
      })),
    };
    return JSON.stringify(state);
  }

  static deserialize(
    serialized: string,
    options: ToolLedgerOptions = {},
  ): ToolLedger {
    const state = parseState(serialized);
    const ledger = new ToolLedger(options);
    ledger.restore(state);
    return ledger;
  }

  static recover(
    serialized: string,
    now: number,
    options: ToolLedgerOptions = {},
  ): { ledger: ToolLedger; outcomes: ToolLedgerRecoveryOutcome[] } {
    const ledger = ToolLedger.deserialize(serialized, options);
    return { ledger, outcomes: ledger.recoverExpired(now) };
  }

  private restore(state: PersistedState): void {
    if (state.entries.length > this.maxEntries) {
      throw malformed("too many retained entries");
    }
    const entryIds = new Set<string>();
    for (const entry of state.entries) {
      validateEntry(entry);
      if (entryIds.has(entry.call_id)) throw malformed("duplicate call id");
      entryIds.add(entry.call_id);
      this.entries.set(entry.call_id, { ...entry });
    }

    const metadataIds = new Set<string>();
    const sequences = new Set<number>();
    for (const metadata of state.metadata) {
      if (
        !validMetadata(metadata) ||
        !entryIds.has(metadata.call_id) ||
        metadataIds.has(metadata.call_id) ||
        sequences.has(metadata.sequence)
      ) {
        throw malformed("invalid call metadata");
      }
      metadataIds.add(metadata.call_id);
      sequences.add(metadata.sequence);
      this.metadata.set(metadata.call_id, { ...metadata });
      this.sequence = Math.max(this.sequence, metadata.sequence + 1);
    }
    if (metadataIds.size !== entryIds.size) {
      throw malformed("every entry must have metadata");
    }

    const taskIds = new Set<string>();
    for (const value of state.tasks) {
      validateTask(value, this.maxRoundsPerTask);
      if (taskIds.has(value.task_id)) throw malformed("duplicate task id");
      taskIds.add(value.task_id);
      const task = emptyTask();
      task.rounds = new Set(value.rounds);
      for (const repetition of value.repetitions) {
        const key = repetitionIdentity(
          repetition.name,
          repetition.canonical_arguments,
        );
        if (task.repetitions.has(key)) {
          throw malformed("duplicate repetition identity");
        }
        task.repetitions.set(key, repetition.count);
      }
      for (const count of value.response_counts) {
        if (task.responseCounts.has(count.key)) {
          throw malformed("duplicate response id");
        }
        task.responseCounts.set(count.key, count.count);
      }
      task.correctiveReaskCallIds = new Set(value.corrective_reask_call_ids);
      if (
        task.correctiveReaskCallIds.size !==
        value.corrective_reask_call_ids.length
      ) {
        throw malformed("duplicate corrective re-ask id");
      }
      this.tasks.set(value.task_id, task);
    }

    for (const metadata of this.metadata.values()) {
      const task = this.tasks.get(metadata.taskId);
      if (!task || !task.rounds.has(metadata.round)) {
        throw malformed("entry references an unknown task or round");
      }
    }
    for (const [taskId, task] of this.tasks) {
      for (const callId of task.correctiveReaskCallIds) {
        if (this.metadata.get(callId)?.taskId !== taskId) {
          throw malformed("corrective re-ask references an unknown call");
        }
      }
    }
  }

  private makeRoom(required: number): boolean {
    if (this.entries.size + required <= this.maxEntries) return true;
    const completed = [...this.entries.values()]
      .filter((entry) => entry.status === "completed")
      .sort(
        (left, right) =>
          this.metadata.get(left.call_id)!.sequence -
          this.metadata.get(right.call_id)!.sequence,
      );
    for (const entry of completed) {
      if (this.entries.size + required <= this.maxEntries) break;
      const taskId = this.metadata.get(entry.call_id)?.taskId;
      this.entries.delete(entry.call_id);
      this.metadata.delete(entry.call_id);
      if (taskId && ![...this.metadata.values()].some((item) => item.taskId === taskId)) {
        this.tasks.delete(taskId);
      }
    }
    return this.entries.size + required <= this.maxEntries;
  }

  private expirePending(now: number): ToolLedgerRecoveryOutcome[] {
    if (!Number.isFinite(now)) {
      throw new ToolLedgerConfigurationError("The ledger clock returned an invalid time.");
    }
    const expired = new Map<string, string[]>();
    for (const entry of this.entries.values()) {
      if (entry.status !== "pending" || entry.issued_at + this.pendingTtlMs > now) {
        continue;
      }
      const taskId = this.metadata.get(entry.call_id)?.taskId;
      if (!taskId) continue;
      expired.set(taskId, [...(expired.get(taskId) ?? []), entry.call_id]);
    }

    const outcomes: ToolLedgerRecoveryOutcome[] = [];
    for (const [taskId, callIds] of expired) {
      for (const [callId, metadata] of this.metadata) {
        if (metadata.taskId === taskId) {
          this.metadata.delete(callId);
          this.entries.delete(callId);
        }
      }
      this.tasks.delete(taskId);
      outcomes.push({
        taskId,
        callIds,
        failure: classifyBridgeFailure("partial_or_unprovable_completion"),
        released: true,
      });
    }
    return outcomes;
  }

  private reject<T>(
    kind: ToolLedgerRejectionKind,
    reason: InternalReason,
  ): ToolLedgerResult<T> {
    return {
      ok: false,
      error: new ToolLedgerError(kind, classifyBridgeFailure(reason)),
    };
  }

  private readNow(): number {
    const value = this.now();
    if (!Number.isFinite(value)) {
      throw new ToolLedgerConfigurationError("The ledger clock returned an invalid time.");
    }
    return value;
  }
}

export function decideToolParallelism(
  metadata: readonly ToolParallelismMetadata[] | undefined,
  maxCalls = MAX_TOOL_CALLS_PER_RESPONSE,
): ToolParallelismDecision {
  if (!metadata || metadata.length === 0) {
    return { parallel: false, reason: "independence_signal_unavailable" };
  }
  if (
    !Number.isSafeInteger(maxCalls) ||
    maxCalls < 1 ||
    metadata.length > maxCalls
  ) {
    return { parallel: false, reason: "call_bound_exceeded" };
  }
  if (metadata.some((call) => !call.requestPermitsParallelCalls)) {
    return { parallel: false, reason: "request_disallows_parallel" };
  }
  if (
    metadata.some(
      (call) =>
        !call.independentlyReadOnly || call.consumesAnotherCallOutput,
    )
  ) {
    return { parallel: false, reason: "mutating_or_dependency_sensitive" };
  }
  return { parallel: true, reason: "allowed" };
}

function canonicalArguments(value: unknown): ToolLedgerResult<string> {
  let parsed = value;
  if (typeof value === "string") {
    try {
      parsed = JSON.parse(value);
    } catch {
      return {
        ok: false,
        error: new ToolLedgerError(
          "invalid_call",
          classifyBridgeFailure("invalid_request"),
        ),
      };
    }
  }
  const result = tryCanonicalizeJson(parsed);
  return result.ok
    ? result
    : {
        ok: false,
        error: new ToolLedgerError(
          "invalid_call",
          classifyBridgeFailure("invalid_request"),
        ),
      };
}

function validateIssueRequest(
  request: IssueCallsRequest,
): ToolLedgerError | null {
  return isRecord(request) &&
    nonEmpty(request.taskId) &&
    nonEmpty(request.responseId) &&
    nonEmpty(request.requestProfileKey) &&
    Array.isArray(request.calls) &&
    request.calls.length > 0 &&
    Number.isSafeInteger(request.round) &&
    request.round > 0
    ? null
    : new ToolLedgerError(
        "invalid_call",
        classifyBridgeFailure("invalid_request"),
      );
}

function emptyTask(): TaskState {
  return {
    rounds: new Set(),
    repetitions: new Map(),
    responseCounts: new Map(),
    correctiveReaskCallIds: new Set(),
  };
}

function repetitionIdentity(name: string, argumentsJson: string): string {
  return JSON.stringify([name, argumentsJson]);
}

function nonEmpty(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function bound(
  name: string,
  value: number,
  minimum = 1,
  maximum = Number.MAX_SAFE_INTEGER,
): void {
  if (!Number.isSafeInteger(value) || value < minimum || value > maximum) {
    throw new ToolLedgerConfigurationError(
      `${name} must be a safe integer between ${minimum} and ${maximum}.`,
    );
  }
}

function parseState(serialized: string): PersistedState {
  if (typeof serialized !== "string") throw new ToolLedgerCodecError(
    "invalid_serialization",
    "The tool ledger state must be a JSON string.",
  );
  let value: unknown;
  try {
    value = JSON.parse(serialized);
  } catch {
    throw new ToolLedgerCodecError("invalid_serialization", "The tool ledger state is not valid JSON.");
  }
  if (!isRecord(value)) throw malformed("the root must be an object");
  if (value.version !== TOOL_LEDGER_STATE_VERSION) {
    throw new ToolLedgerCodecError("unsupported_version", "The tool ledger state version is not supported.");
  }
  if (
    !Array.isArray(value.entries) ||
    !Array.isArray(value.metadata) ||
    !Array.isArray(value.tasks)
  ) {
    throw malformed("entries, metadata, and tasks must be arrays");
  }
  return {
    version: value.version as number,
    entries: value.entries as IssuedCall[],
    metadata: value.metadata as PersistedState["metadata"],
    tasks: value.tasks as PersistedState["tasks"],
  };
}

function validateEntry(value: IssuedCall): void {
  if (
    !isRecord(value) ||
    !nonEmpty(value.call_id) ||
    !nonEmpty(value.name) ||
    !nonEmpty(value.type) ||
    !nonEmpty(value.canonical_arguments) ||
    !nonEmpty(value.request_profile_key) ||
    !Number.isFinite(value.issued_at) ||
    (value.status !== "pending" && value.status !== "completed")
  ) {
    throw malformed("invalid issued call");
  }
  if (!isCanonical(value.canonical_arguments)) {
    throw malformed("non-canonical arguments");
  }
}

function validMetadata(
  value: CallMetadata & { call_id: string },
): boolean {
  return (
    isRecord(value) &&
    nonEmpty(value.call_id) &&
    nonEmpty(value.taskId) &&
    nonEmpty(value.responseId) &&
    Number.isSafeInteger(value.round) &&
    value.round > 0 &&
    Number.isSafeInteger(value.sequence) &&
    value.sequence >= 0
  );
}

function validateTask(
  value: PersistedState["tasks"][number],
  maxRounds: number,
): void {
  if (
    !isRecord(value) ||
    !nonEmpty(value.task_id) ||
    !Array.isArray(value.rounds) ||
    !Array.isArray(value.repetitions) ||
    !Array.isArray(value.response_counts) ||
    !Array.isArray(value.corrective_reask_call_ids) ||
    value.rounds.length > maxRounds ||
    value.repetitions.length > MAX_TOOL_LEDGER_ENTRIES
  ) {
    throw malformed("invalid task state");
  }
  if (new Set(value.rounds).size !== value.rounds.length ||
    value.rounds.some((round) => !Number.isSafeInteger(round) || round < 1 || round > maxRounds)) {
    throw malformed("invalid task rounds");
  }
  for (const repetition of value.repetitions) {
    if (
      !isRecord(repetition) ||
      !nonEmpty(repetition.name) ||
      !nonEmpty(repetition.canonical_arguments) ||
      !Number.isSafeInteger(repetition.count) ||
      repetition.count < 1 ||
      repetition.count > MAX_TOOL_REPETITIONS_PER_TASK
    ) {
      throw malformed("invalid repetition state");
    }
    if (!isCanonical(repetition.canonical_arguments)) {
      throw malformed("non-canonical repetition arguments");
    }
  }
  for (const count of value.response_counts) {
    if (
      !isRecord(count) ||
      !nonEmpty(count.key) ||
      !Number.isSafeInteger(count.count) ||
      count.count < 1 ||
      count.count > MAX_TOOL_CALLS_PER_RESPONSE
    ) {
      throw malformed("invalid response count");
    }
  }
  if (
    value.corrective_reask_call_ids.length > MAX_TOOL_LEDGER_ENTRIES ||
    new Set(value.corrective_reask_call_ids).size !==
      value.corrective_reask_call_ids.length ||
    value.corrective_reask_call_ids.some((callId) => !nonEmpty(callId))
  ) {
    throw malformed("invalid corrective re-ask state");
  }
}

function isCanonical(value: string): boolean {
  try {
    const result = tryCanonicalizeJson(JSON.parse(value));
    return result.ok && result.value === value;
  } catch {
    return false;
  }
}

function malformed(message: string): ToolLedgerCodecError {
  return new ToolLedgerCodecError("malformed_state", message);
}
