import { randomUUID } from "node:crypto";
import WebSocket from "ws";
import { getTokenPath, loadToken } from "../cli/token-helpers";
import {
  tryBuildAssistantResponseFromChatCompletionPayload,
  tryExtractSimulatedResponsePayload,
} from "./openai";
import { OpenAiTransformModes } from "./types";
import type {
  ChatResult,
  CreateConversationResult,
  JsonObject,
  JsonValue,
  ParsedOpenAiRequest,
  SubstrateOptions,
  SubstrateStreamUpdate,
  WrapperOptions,
} from "./types";
import {
  computeTrailingDelta,
  extractBearerToken,
  isJsonObject,
  tryGetString,
  tryParseJsonObject,
  tryReadJwtPayload,
  extractGraphErrorMessage,
  tryGetInt,
} from "./utils";
import { DebugMarkdownLogger } from "./logger";
import {
  SubstrateSessionStore,
  TurnQueueWaitError,
} from "./substrate-session-store";
import { ConcurrencyLimiter } from "./concurrency";
import { resolveSubstrateCapabilities } from "./substrate-capabilities";
import { uploadSubstrateImages } from "./substrate-image-upload";
import { classifyBridgeFailure } from "./failure-classifier";
import type { BridgeObservability } from "./observability";
export { resolveSubstrateTone } from "./substrate-capabilities";

export class CopilotGraphClient {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
  ) {}

  async createConversation(
    authorizationHeader: string,
  ): Promise<CreateConversationResult> {
    const uri = this.buildAbsoluteUri(this.options.createConversationPath);
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "application/json",
      "Content-Type": "application/json",
    });
    const body = "{}";

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    const rawBody = await response.text();
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      rawBody,
      true,
    );

    const json = tryParseJsonObject(rawBody);
    return {
      isSuccess: response.ok,
      statusCode: response.status,
      conversationId: tryGetString(json, "id"),
      rawBody,
    };
  }

  async chat(
    authorizationHeader: string,
    conversationId: string,
    payload: JsonObject,
  ): Promise<ChatResult> {
    const uri = this.buildAbsoluteUri(
      resolveConversationPath(this.options.chatPathTemplate, conversationId),
    );
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "application/json",
      "Content-Type": "application/json",
    });
    const body = JSON.stringify(payload);

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    const rawBody = await response.text();
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      rawBody,
      true,
    );

    return {
      isSuccess: response.ok,
      statusCode: response.status,
      responseJson: tryParseJsonObject(rawBody),
      rawBody,
      assistantText: null,
      conversationId: null,
      upstreamRequestPayload: payload,
      upstreamResponsePayload: parseRawJsonValue(rawBody),
    };
  }

  async chatOverStream(
    authorizationHeader: string,
    conversationId: string,
    payload: JsonObject,
  ): Promise<Response> {
    const uri = this.buildAbsoluteUri(
      resolveConversationPath(
        this.options.chatOverStreamPathTemplate,
        conversationId,
      ),
    );
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "text/event-stream",
      "Content-Type": "application/json",
    });
    const body = JSON.stringify(payload);

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      null,
      false,
    );
    return response;
  }

  private buildAbsoluteUri(relativePath: string): URL {
    try {
      return new URL(relativePath);
    } catch {
      const baseUrl =
        this.options.graphBaseUrl?.trim() || "https://graph.microsoft.com";
      const normalized = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
      return new URL(relativePath.replace(/^\/+/, ""), normalized);
    }
  }
}

export type SubstrateWebSocketConnector = (
  url: URL,
  origin: string | undefined,
  timeoutMs: number,
  keepAliveMs: number,
) => Promise<WebSocket>;

export type SubstrateWebSocketReceiver = {
  next: (timeoutMs: number) => Promise<string | null>;
  dispose: () => void;
};

export type SubstrateReceiverFactory = (
  ws: WebSocket,
) => SubstrateWebSocketReceiver;

// SignalR Hub Protocol types 1/2/3/6/7 are the observed Substrate decoder
// allowlist; keep private-protocol constants beside this decoder (PRD §15
// rule 5). https://github.com/dotnet/aspnetcore/blob/main/src/SignalR/docs/specs/HubProtocol.md
const SignalRInvocationType = 1;
const SignalRStreamItemType = 2;
const SignalRCompletionType = 3;
const SignalRPingType = 6;
const SignalRCloseType = 7;
const AllowlistedSubstrateFrameTypes = new Set<number>([
  SignalRInvocationType,
  SignalRStreamItemType,
  SignalRCompletionType,
  SignalRPingType,
  SignalRCloseType,
]);
const KnownSubstrateMessageTypes = new Set(["chat", "disengaged"]);
const MaxSubstrateDriftObservations = 8;
type SubstrateDriftReason =
  | "unknown_frame_type"
  | "unknown_message_type"
  | "malformed_semantic_payload"
  | "unrecognized_terminal_semantics";

export class CopilotSubstrateClient {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
    private readonly sessionStore = new SubstrateSessionStore(
      options.conversationTtlMinutes,
    ),
    private readonly connect: SubstrateWebSocketConnector = connectWebSocket,
    private readonly createReceiver: SubstrateReceiverFactory = createWebSocketReceiver,
    private readonly concurrencyLimiter: ConcurrencyLimiter = new ConcurrencyLimiter(
      options.substrate.concurrencyLimit ?? 0,
      options.substrate.acquireTimeoutMs ?? 0,
    ),
    private readonly observability: BridgeObservability | null = null,
  ) {}

  createConversation(): CreateConversationResult {
    return {
      isSuccess: true,
      statusCode: 200,
      conversationId: randomUUID(),
      rawBody: "",
    };
  }

  async chat(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate?: (update: SubstrateStreamUpdate) => Promise<void>,
    signal: AbortSignal | null = null,
    taskDeadlineMs: number | null = null,
  ): Promise<ChatResult> {
    return this.chatCore(
      authorizationHeader,
      conversationId,
      request,
      isStartOfSession,
      onStreamUpdate ?? null,
      signal,
      taskDeadlineMs,
    );
  }

  async chatStream(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate: (update: SubstrateStreamUpdate) => Promise<void>,
    signal: AbortSignal | null = null,
    taskDeadlineMs: number | null = null,
  ): Promise<ChatResult> {
    return this.chatCore(
      authorizationHeader,
      conversationId,
      request,
      isStartOfSession,
      onStreamUpdate,
      signal,
      taskDeadlineMs,
    );
  }

  private async chatCore(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate: ((update: SubstrateStreamUpdate) => Promise<void>) | null,
    signal: AbortSignal | null = null,
    taskDeadlineMs: number | null = null,
  ): Promise<ChatResult> {
    const sessionId = this.sessionStore.getOrCreate(conversationId);
    const startedAtMs = Date.now();
    const queueDepth =
      this.sessionStore.queuedTurnCount + this.concurrencyLimiter.queueDepth;
    try {
      const result = await this.sessionStore.runExclusive(
        sessionId,
        async () => {
          const releaseSlot = await this.concurrencyLimiter.acquire(
            signal,
            taskDeadlineMs,
          );
          try {
            return await this.chatCoreUnlocked(
              authorizationHeader,
              conversationId,
              sessionId,
              request,
              isStartOfSession,
              onStreamUpdate,
              signal,
              taskDeadlineMs,
            );
          } finally {
            releaseSlot();
          }
        },
        signal,
        taskDeadlineMs,
      );
      this.observability?.record("queue_wait", {
        depth: queueDepth,
        durationMs: Math.max(0, Date.now() - startedAtMs),
        outcome: "acquired",
      });
      this.observability?.record("substrate_terminal", {
        success: result.isSuccess,
        statusCode: result.statusCode,
        latencyMs: Math.max(0, Date.now() - startedAtMs),
      });
      if (result.statusCode === 499) {
        this.observability?.record("cancellation", { phase: "substrate" });
      }
      return result;
    } catch (error) {
      if (error instanceof TurnQueueWaitError) {
        this.observability?.record("queue_wait", {
          depth: queueDepth,
          durationMs: Math.max(0, Date.now() - startedAtMs),
          outcome: error.reason,
        });
        if (error.reason === "cancelled") {
          this.observability?.record("cancellation", { phase: "queue" });
        }
        return buildFailure(
          error.reason === "cancelled"
            ? 499
            : error.reason === "deadline"
              ? 504
              : 429,
          error.message,
        );
      }
      throw error;
    }
  }

  private async chatCoreUnlocked(
    authorizationHeader: string,
    conversationId: string,
    sessionId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate: ((update: SubstrateStreamUpdate) => Promise<void>) | null,
    signal: AbortSignal | null = null,
    taskDeadlineMs: number | null = null,
  ): Promise<ChatResult> {
    if (signal?.aborted) {
      return buildFailure(499, "Substrate turn cancelled before start.");
    }
    const rawToken = extractBearerToken(authorizationHeader);
    if (!rawToken) {
      return buildFailure(401, "Missing Bearer token.");
    }

    const tokenPayload = tryReadJwtPayload(rawToken);
    let objectId = tokenPayload ? tryGetString(tokenPayload, "oid") : null;
    let tenantId = tokenPayload ? tryGetString(tokenPayload, "tid") : null;
    if (!objectId || !tenantId) {
      const tokenPath = await getTokenPath();
      const tokenState = await loadToken(tokenPath);
      if (tokenState?.token === rawToken) {
        objectId = tokenState.oid?.trim() || objectId;
        tenantId = tokenState.tid?.trim() || tenantId;
      }
    }
    if (!objectId || !tenantId) {
      return buildFailure(
        400,
        "Authorization token must be a JWT or have cached oid/tid metadata so Substrate can be addressed correctly.",
      );
    }

    let resolvedConversationId = conversationId;
    let imageAnnotations: JsonObject[] = [];
    if ((request.images ?? []).length > 0) {
      const uploadResult = await uploadSubstrateImages(request.images ?? [], {
        rawToken,
        objectId,
        tenantId,
        conversationId: resolvedConversationId,
        substrate: this.options.substrate,
        signal,
      });
      if (!uploadResult.ok) {
        return buildFailure(
          uploadResult.aborted || signal?.aborted ? 499 : 502,
          uploadResult.aborted || signal?.aborted
            ? "Substrate image upload was cancelled."
            : "Substrate image upload failed.",
          null,
          null,
          uploadResult.aborted || signal?.aborted
            ? null
            : classifyBridgeFailure("image_upload_failed").code,
        );
      }
      imageAnnotations = uploadResult.annotations;
    }

    const clientRequestId = randomUUID();
    const requestUri = buildSubstrateHubUri(
      this.options,
      objectId,
      tenantId,
      rawToken,
      clientRequestId,
      sessionId,
      resolvedConversationId,
    );
    const {
      invocationTimeoutMs: configuredTimeoutMs,
      handshakeTimeoutMs: configuredHandshakeTimeoutMs,
      turnDeadlineMs: configuredTurnDeadlineMs,
    } =
      resolveSubstrateDeadlines(this.options.substrate, Date.now());
    const effectiveTaskDeadlineMs =
      taskDeadlineMs === null
        ? Number.POSITIVE_INFINITY
        : taskDeadlineMs;
    const remainingTaskMs = effectiveTaskDeadlineMs - Date.now();
    if (remainingTaskMs <= 0) {
      return buildFailure(504, "M365 task-level deadline exceeded.");
    }
    const timeoutMs = Math.min(configuredTimeoutMs, remainingTaskMs);
    const handshakeTimeoutMs = Math.min(
      configuredHandshakeTimeoutMs,
      remainingTaskMs,
    );
    const turnDeadlineMs = Math.min(
      configuredTurnDeadlineMs,
      effectiveTaskDeadlineMs,
    );
    const ws = await connectWithAbort(
      this.connect(
        requestUri,
        this.options.substrate.origin ?? undefined,
        timeoutMs,
        this.options.substrate.keepAliveSeconds > 0
          ? this.options.substrate.keepAliveSeconds * 1000
          : 15_000,
      ),
      signal,
    );
    if (ws instanceof Error) {
      if (ws.name === "AbortError") {
        return buildFailure(
          499,
          "Substrate turn cancelled during websocket connection.",
        );
      }
      return buildFailure(
        502,
        `Substrate websocket request failed. ${ws.message}`,
      );
    }
    const remainingHandshakeMs = Math.min(
      handshakeTimeoutMs,
      turnDeadlineMs - Date.now(),
      effectiveTaskDeadlineMs - Date.now(),
    );
    if (remainingHandshakeMs <= 0) {
      try {
        ws.close(1000, "task deadline exceeded");
      } catch {
        // ignore
      }
      return buildFailure(504, "M365 task-level deadline exceeded.");
    }

    const transcript: string[] = [];
    let invocationPayload: JsonObject | null = null;
    let activeInvocationId: string | null = null;
    let invocationStarted = false;
    let invocationSendPromise: Promise<void> | null = null;
    let cancellationRequested = false;
    let cancellationStarted = false;
    let socketClosed = false;
    let receiverDisposed = false;
    const receiver = this.createReceiver(ws);
    let providerDriftObserved = false;
    const driftObservationKeys = new Set<string>();
    const recordDrift = (
      reason: SubstrateDriftReason,
      frameType: number | null,
      envelope: JsonObject | null,
      target: string | null = null,
      messageType: string | null = null,
    ) => {
      providerDriftObserved = true;
      if (driftObservationKeys.size >= MaxSubstrateDriftObservations) {
        return;
      }
      const observation = buildSubstrateDriftObservation(
        reason,
        frameType,
        envelope,
        target,
        messageType,
      );
      const key = JSON.stringify(observation);
      if (driftObservationKeys.has(key)) {
        return;
      }
      driftObservationKeys.add(key);
      this.observability?.record(
        "provider_drift",
        observation,
        clientRequestId,
      );
    };
    const disposeReceiver = () => {
      if (receiverDisposed) {
        return;
      }
      receiverDisposed = true;
      receiver.dispose();
    };
    const closeSocket = (reason: string | null, hard: boolean) => {
      if (socketClosed) {
        return;
      }
      socketClosed = true;
      try {
        if (
          hard &&
          typeof (ws as { terminate?: () => void }).terminate === "function"
        ) {
          (ws as { terminate: () => void }).terminate();
          return;
        }
        if (reason) {
          ws.close(1000, reason);
        } else {
          ws.close();
        }
      } catch {
        // ignore
      }
    };
    const cancelUpstream = async (): Promise<void> => {
      let stopSent = false;
      let acknowledgement: SubstrateCancellationAcknowledgement | null = null;
      try {
        if (!invocationStarted && invocationSendPromise) {
          const invocationSend = await settlePromiseWithin(
            invocationSendPromise,
            SubstrateCancelAckTimeoutMs,
          );
          if (invocationSend.completed && !invocationSend.error) {
            invocationStarted = true;
          } else {
            invocationSendPromise = null;
          }
        }

        const invocationId = activeInvocationId;
        if (
          invocationStarted &&
          invocationId &&
          ws.readyState === WebSocket.OPEN
        ) {
          acknowledgement = createSubstrateCancellationAcknowledgement(
            ws,
            invocationId,
            SubstrateCancelAckTimeoutMs,
          );
          const stop = await settlePromiseWithin(
            sendFrame(ws, requestUri, this.logger, {
              type: SignalRCancelInvocationType,
              invocationId,
            }),
            SubstrateCancelAckTimeoutMs,
          );
          stopSent = stop.completed && !stop.error;
        }
      } catch {
        // Cancellation is best effort and must never reject the turn.
      }

      disposeReceiver();
      if (stopSent && acknowledgement) {
        await acknowledgement.wait;
      }
      acknowledgement?.dispose();
      closeSocket("client disconnected", true);
    };
    const onAbort = () => {
      cancellationRequested = true;
      if (cancellationStarted) {
        return;
      }
      cancellationStarted = true;
      void cancelUpstream().catch(() => {
        disposeReceiver();
        closeSocket("client disconnected", true);
      });
    };
    if (signal) {
      if (signal.aborted) {
        onAbort();
      } else {
        signal.addEventListener("abort", onAbort, { once: true });
      }
    }
    try {
      if (cancellationRequested || signal?.aborted) {
        return buildFailure(
          499,
          "Substrate turn cancelled by client disconnect.",
        );
      }
      await sendFrame(ws, requestUri, this.logger, {
        protocol: "json",
        version: 1,
      });
      const handshakeDeadlineMs = Date.now() + remainingHandshakeMs;
      const handshakePayload = await receiver.next(remainingHandshakeMs);
      if (handshakePayload === null) {
        if (cancellationRequested || signal?.aborted) {
          return buildFailure(
            499,
            "Substrate turn cancelled by client disconnect.",
          );
        }
        if (Date.now() >= handshakeDeadlineMs) {
          return buildFailure(504, "Substrate websocket handshake timed out.");
        }
        return buildFailure(
          502,
          "Substrate websocket closed during handshake.",
        );
      }
      await this.logger.logSubstrateFrame(
        requestUri.toString(),
        "response",
        handshakePayload,
      );
      transcript.push(handshakePayload);
      if (onStreamUpdate) {
        await onStreamUpdate({
          deltaText: null,
          conversationId: resolvedConversationId,
          upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
        });
      }

      for (const frame of splitFrames(handshakePayload)) {
        const frameJson = tryParseJsonObject(frame);
        const handshakeError = tryGetString(frameJson, "error");
        if (handshakeError) {
          return buildFailure(
            502,
            `Substrate handshake failed. ${handshakeError}`,
          );
        }
      }

      if (cancellationRequested || signal?.aborted) {
        return buildFailure(
            499,
            "Substrate turn cancelled by client disconnect.",
        );
      }
      await sendFrame(ws, requestUri, this.logger, { type: 6 });
      invocationPayload = buildInvocationPayload(
        request,
        resolvedConversationId,
        sessionId,
        clientRequestId,
        isStartOfSession,
        this.options,
        this.observability,
        imageAnnotations,
      );
      activeInvocationId = tryGetString(invocationPayload, "invocationId");
      if (onStreamUpdate) {
        await onStreamUpdate({
          deltaText: null,
          conversationId: resolvedConversationId,
          upstreamRequestPayload: invocationPayload,
          upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
        });
      }
      if (cancellationRequested || signal?.aborted) {
        return buildFailure(
          499,
          "Substrate turn cancelled by client disconnect.",
        );
      }
      invocationSendPromise = sendFrame(
        ws,
        requestUri,
        this.logger,
        invocationPayload,
      );
      try {
        await invocationSendPromise;
        invocationStarted = true;
      } finally {
        invocationSendPromise = null;
      }

      let assistantText = "";
      let deltaBuilder = "";
      let responseError: string | null = null;
      let responseErrorCode: string | null = null;
      let completed = false;
      let receivedSuccessfulTerminal = false;
      let activeBotMessageId: string | null = null;
      const foldedFrames = new Set<string>();
      const logDeltaSource = async (
        source: "writeAtCursor" | "snapshotText",
        deltaText: string,
      ) => {
        await this.logger.logSubstrateFrame(
          requestUri.toString(),
          "delta",
          JSON.stringify({
            source,
            conversationId: resolvedConversationId,
            deltaLength: deltaText.length,
            deltaPreview: deltaText.slice(0, 120),
          }),
        );
      };

      while (!completed && ws.readyState === WebSocket.OPEN) {
        if (cancellationRequested || signal?.aborted) {
          break;
        }
        const remainingMs = turnDeadlineMs - Date.now();
        if (remainingMs <= 0) {
          return buildFailure(504, "Substrate websocket turn timed out.");
        }
        const payload = await receiver.next(remainingMs);
        if (payload === null) {
          if (cancellationRequested || signal?.aborted) {
            break;
          }
          if (Date.now() >= turnDeadlineMs) {
            return buildFailure(
              504,
              "Substrate websocket turn timed out.",
              invocationPayload,
              buildSubstrateTranscriptPayload(transcript),
            );
          }
          break;
        }
        if (cancellationRequested || signal?.aborted) {
          break;
        }
        await this.logger.logSubstrateFrame(
          requestUri.toString(),
          "response",
          payload,
        );
        if (cancellationRequested || signal?.aborted) {
          break;
        }
        transcript.push(payload);
        if (onStreamUpdate) {
          await onStreamUpdate({
            deltaText: null,
            conversationId: resolvedConversationId,
            upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
          });
        }
        if (cancellationRequested || signal?.aborted) {
          break;
        }

        for (const frame of splitFrames(payload)) {
          if (cancellationRequested || signal?.aborted) {
            break;
          }
          if (!frame.trim()) {
            continue;
          }
          const json = tryParseJsonObject(frame);
          if (!json) {
            recordDrift("malformed_semantic_payload", null, null);
            continue;
          }
          const frameType = tryGetInt(json, "type");
          if (frameType === SignalRPingType) {
            await sendFrame(ws, requestUri, this.logger, { type: 6 });
            continue;
          }
          if (
            frameType === null ||
            !AllowlistedSubstrateFrameTypes.has(frameType)
          ) {
            recordDrift("unknown_frame_type", frameType, json);
            continue;
          }

          const frameIdentity = buildSubstrateFrameIdentity(json, frame.trim());
          if (foldedFrames.has(frameIdentity)) {
            continue;
          }
          foldedFrames.add(frameIdentity);

          const frameRequestId = extractSubstrateRequestId(json);
          if (frameRequestId && frameRequestId !== clientRequestId) {
            continue;
          }

          const frameBotMessageId = extractSubstrateBotMessageId(json);
          if (
            !frameRequestId &&
            activeBotMessageId &&
            frameBotMessageId &&
            frameBotMessageId !== activeBotMessageId
          ) {
            continue;
          }

          if (frameRequestId === clientRequestId && frameBotMessageId) {
            activeBotMessageId = frameBotMessageId;
          } else if (!activeBotMessageId && frameBotMessageId) {
            activeBotMessageId = frameBotMessageId;
          }

          if (
            frameType === SignalRInvocationType ||
            frameType === SignalRStreamItemType
          ) {
            for (const issue of inspectSubstrateSemanticPayload(json)) {
              recordDrift(
                issue.reason,
                frameType,
                json,
                issue.target,
                issue.messageType,
              );
            }
          }

          const extractedConversationId = extractSubstrateConversationId(json);
          if (
            extractedConversationId &&
            extractedConversationId !== resolvedConversationId
          ) {
            resolvedConversationId = extractedConversationId;
            this.sessionStore.set(resolvedConversationId, sessionId);
            if (onStreamUpdate) {
              await onStreamUpdate({
                deltaText: null,
                conversationId: resolvedConversationId,
              });
            }
            if (cancellationRequested || signal?.aborted) {
              break;
            }
          }

          const deltaText = extractSubstrateDeltaText(json);
          if (deltaText) {
            deltaBuilder += deltaText;
            await logDeltaSource("writeAtCursor", deltaText);
            if (cancellationRequested || signal?.aborted) {
              break;
            }
            if (onStreamUpdate) {
              await onStreamUpdate({
                deltaText,
                conversationId: resolvedConversationId,
              });
            }
            if (cancellationRequested || signal?.aborted) {
              break;
            }
          }

          const disengagedMessage = extractSubstrateDisengagedMessage(json);
          if (disengagedMessage) {
            responseError = disengagedMessage;
            responseErrorCode = "substrate_disengaged";
            completed = true;
            break;
          }

          const extractedAssistantText = extractSubstrateAssistantText(json);
          if (extractedAssistantText) {
            assistantText = extractedAssistantText;
            const snapshotDelta = computeTrailingDelta(
              deltaBuilder,
              extractedAssistantText,
            );
            if (snapshotDelta) {
              deltaBuilder += snapshotDelta;
              await logDeltaSource("snapshotText", snapshotDelta);
              if (cancellationRequested || signal?.aborted) {
                break;
              }
              if (onStreamUpdate) {
                await onStreamUpdate({
                  deltaText: snapshotDelta,
                  conversationId: resolvedConversationId,
                });
              }
              if (cancellationRequested || signal?.aborted) {
                break;
              }

            }

            if (
              request.transformMode === OpenAiTransformModes.Simulated &&
              this.options.substrate.earlyCompleteOnSimulatedPayload &&
              resolveCompleteSimulatedAssistantText(
                extractedAssistantText,
                deltaBuilder,
              )
            ) {
              assistantText =
                resolveCompleteSimulatedAssistantText(
                  extractedAssistantText,
                  deltaBuilder,
                ) ?? assistantText;
              completed = true;
              break;
            }
          }

          const frameError = tryGetString(json, "error");
          if (frameError) {
            responseError = frameError;
          }
          const resultValue = extractSubstrateResultValue(json);
          if (resultValue && !isSubstrateResultSuccess(resultValue)) {
            responseError =
              extractSubstrateResultMessage(json) ??
              `Substrate returned result '${resultValue}'.`;
          }

          if (frameType === SignalRCloseType) {
            const hasUsableTerminalError =
              Boolean(tryGetString(json, "error")?.trim()) ||
              Boolean(
                resultValue &&
                  !isSubstrateResultSuccess(resultValue) &&
                  extractSubstrateResultMessage(json)?.trim(),
              );
            if (!hasUsableTerminalError) {
              recordDrift(
                "unrecognized_terminal_semantics",
                frameType,
                json,
              );
              responseError =
                "Substrate websocket closed with an unrecognized terminal frame.";
              responseErrorCode = "provider_drift";
            } else {
              responseError ??=
                "Substrate websocket closed with an error frame.";
              responseErrorCode ??= "substrate_terminal_error";
            }
            completed = true;
            break;
          }

          if (frameType === SignalRCompletionType) {
            if (
              hasSubstrateResultValue(json) &&
              (!resultValue || !isSubstrateResultSuccess(resultValue))
            ) {
              recordDrift(
                "unrecognized_terminal_semantics",
                frameType,
                json,
              );
              responseError =
                "Substrate completion returned an unrecognized result.";
              responseErrorCode = "provider_drift";
            }
            if (!responseError) {
              receivedSuccessfulTerminal = true;
            }
            completed = true;
            break;
          }
        }
      }

      if (!cancellationRequested && !signal?.aborted) {
        closeSocket("completed", false);
      }

      if (cancellationRequested || signal?.aborted) {
        return buildFailure(
          499,
          "Substrate turn cancelled by client disconnect.",
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
        );
      }

      // SignalR type 7/error frames are terminal failures. Never turn a
      // partially streamed message into a successful Codex turn: that would
      // make a failed write or patch look completed to the client.
      if (responseError) {
        return buildFailure(
          502,
          `Substrate chat failed. ${responseError}`,
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
          providerDriftObserved ? "provider_drift" : responseErrorCode,
        );
      }

      if (!receivedSuccessfulTerminal) {
        return buildFailure(
          502,
          providerDriftObserved
            ? "Substrate chat ended after provider protocol drift."
            : "Substrate chat ended without a successful terminal frame.",
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
          providerDriftObserved
            ? "provider_drift"
            : "substrate_incomplete_terminal",
        );
      }

      if (!assistantText && deltaBuilder) {
        assistantText = deltaBuilder;
      }

      if (!assistantText) {
        return buildFailure(
          502,
          "Substrate chat returned no assistant content.",
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
          providerDriftObserved ? "provider_drift" : null,
        );
      }

      this.sessionStore.set(resolvedConversationId, sessionId);

      return {
        isSuccess: true,
        statusCode: 200,
        responseJson: buildNormalizedConversation(
          resolvedConversationId,
          request.promptText,
          assistantText,
        ),
        rawBody: transcript.join("\n"),
        assistantText,
        conversationId: resolvedConversationId,
        upstreamRequestPayload: invocationPayload,
        upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
      };
    } catch (error) {
      if (cancellationRequested || signal?.aborted) {
        return buildFailure(
          499,
          "Substrate turn cancelled by client disconnect.",
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
        );
      }
      const message = String(error);
      if (message.toLowerCase().includes("timeout")) {
        return buildFailure(
          504,
          "Substrate websocket request timed out.",
          invocationPayload,
          buildSubstrateTranscriptPayload(transcript),
        );
      }
      return buildFailure(
        502,
        `Unexpected Substrate websocket failure. ${message}`,
        invocationPayload,
        buildSubstrateTranscriptPayload(transcript),
      );
    } finally {
      if (signal) {
        signal.removeEventListener("abort", onAbort);
      }
      disposeReceiver();
      if (!cancellationRequested && !signal?.aborted) {
        closeSocket(null, false);
      }
    }
  }
}

function resolveConversationPath(
  pathTemplate: string,
  conversationId: string,
): string {
  const template = pathTemplate?.trim()
    ? pathTemplate
    : "/beta/copilot/conversations/{conversationId}/chat";
  return template.replaceAll(
    "{conversationId}",
    encodeURIComponent(conversationId),
  );
}

type SubstrateSemanticIssue = {
  reason: SubstrateDriftReason;
  target: string | null;
  messageType: string | null;
};

function inspectSubstrateSemanticPayload(
  envelope: JsonObject,
): SubstrateSemanticIssue[] {
  const issues: SubstrateSemanticIssue[] = [];
  const addIssue = (
    reason: SubstrateDriftReason,
    target: string | null = null,
    messageType: string | null = null,
  ) => {
    issues.push({ reason, target, messageType });
  };
  const target = envelope.target;
  if (target !== undefined && typeof target !== "string") {
    addIssue("malformed_semantic_payload");
  } else if (
    typeof target === "string" &&
    target.trim() &&
    target.trim().toLowerCase() !== "update"
  ) {
    addIssue("unknown_message_type", target.trim());
  }

  const inspectMessageArray = (value: JsonValue | undefined) => {
    if (value === undefined) {
      return;
    }
    if (!Array.isArray(value)) {
      addIssue("malformed_semantic_payload");
      return;
    }
    for (const message of value) {
      if (!isJsonObject(message)) {
        addIssue("malformed_semantic_payload");
        continue;
      }
      const messageType = message.messageType;
      if (messageType !== undefined && typeof messageType !== "string") {
        addIssue("malformed_semantic_payload");
      } else if (
        typeof messageType === "string" &&
        messageType.trim() &&
        !KnownSubstrateMessageTypes.has(messageType.trim().toLowerCase())
      ) {
        addIssue("unknown_message_type", null, messageType.trim());
      }
      for (const textKey of ["text", "hiddenText", "spokenText"]) {
        if (
          Object.prototype.hasOwnProperty.call(message, textKey) &&
          typeof message[textKey] !== "string"
        ) {
          addIssue("malformed_semantic_payload");
        }
      }
    }
  };
  const inspectContainer = (container: JsonObject) => {
    inspectMessageArray(container.messages);
    if (
      Object.prototype.hasOwnProperty.call(container, "writeAtCursor") &&
      typeof container.writeAtCursor !== "string"
    ) {
      addIssue("malformed_semantic_payload");
    }
    if (container.item !== undefined) {
      if (!isJsonObject(container.item)) {
        addIssue("malformed_semantic_payload");
      } else {
        inspectMessageArray(container.item.messages);
      }
    }
  };

  inspectContainer(envelope);
  if (envelope.arguments !== undefined) {
    if (!Array.isArray(envelope.arguments)) {
      addIssue("malformed_semantic_payload");
    } else {
      for (const argument of envelope.arguments) {
        if (!isJsonObject(argument)) {
          addIssue("malformed_semantic_payload");
          continue;
        }
        inspectContainer(argument);
      }
    }
  }
  return issues;
}

function buildSubstrateDriftObservation(
  reason: SubstrateDriftReason,
  frameType: number | null,
  envelope: JsonObject | null,
  target: string | null,
  messageType: string | null,
): JsonObject {
  const observation: JsonObject = {
    reason,
    frameType,
    shape: describeSubstrateFrameShape(envelope),
  };
  const safeTarget = sanitizeSubstrateDiagnosticLabel(target);
  const safeMessageType = sanitizeSubstrateDiagnosticLabel(messageType);
  if (safeTarget) {
    observation.target = safeTarget;
  }
  if (safeMessageType) {
    observation.messageType = safeMessageType;
  }
  return observation;
}

function describeSubstrateFrameShape(envelope: JsonObject | null): string {
  if (!envelope) {
    return "unparseable";
  }
  const keys = Object.keys(envelope)
    .sort()
    .map((key) => key.replace(/[^A-Za-z0-9_.-]/g, "_").slice(0, 24))
    .filter(Boolean);
  if (keys.length === 0) {
    return "empty";
  }
  const shown = keys.slice(0, 8);
  return keys.length > shown.length
    ? `${shown.join(",")}+${keys.length - shown.length}`
    : shown.join(",");
}

function sanitizeSubstrateDiagnosticLabel(value: string | null): string | null {
  if (!value) {
    return null;
  }
  const normalized = value.trim();
  if (!normalized) {
    return null;
  }
  if (
    /(?:prompt|assistant|token|secret|bearer|cookie|account|tenant)/i.test(
      normalized,
    )
  ) {
    return "[redacted-label]";
  }
  if (!/^[A-Za-z][A-Za-z0-9_.:-]{0,31}$/.test(normalized)) {
    return "present";
  }
  return normalized;
}

function buildFailure(
  statusCode: number,
  message: string,
  upstreamRequestPayload: JsonValue | null = null,
  upstreamResponsePayload: JsonValue | null = null,
  errorCode: string | null = null,
): ChatResult {
  return {
    isSuccess: false,
    statusCode,
    responseJson: null,
    rawBody: JSON.stringify({ message, code: errorCode }),
    assistantText: null,
    conversationId: null,
    errorCode,
    upstreamRequestPayload,
    upstreamResponsePayload,
  };
}

export function buildSubstrateHubUri(
  options: WrapperOptions,
  objectId: string,
  tenantId: string,
  accessToken: string,
  clientRequestId: string,
  sessionId: string,
  conversationId: string,
): URL {
  let baseHub =
    options.substrate.hubPath?.trim() ||
    "wss://substrate.office.com/m365Copilot/Chathub/";
  if (!baseHub.endsWith("/")) {
    baseHub += "/";
  }
  const hubUri = new URL(
    `${encodeURIComponent(objectId)}@${encodeURIComponent(tenantId)}`,
    baseHub,
  );

  const query = new URLSearchParams({
    chatsessionid: randomUUID().replaceAll("-", ""),
    XRoutingParameterSessionKey: randomUUID().replaceAll("-", ""),
    ClientRequestId: clientRequestId,
    "X-SessionId": sessionId,
    ConversationId: conversationId,
    access_token: accessToken,
  });

  if (options.substrate.source?.trim()) {
    const sourceValue = options.substrate.quoteSourceInQuery
      ? `"${options.substrate.source}"`
      : options.substrate.source;
    query.set("source", sourceValue);
  }
  if (options.substrate.scenario?.trim()) {
    query.set("scenario", options.substrate.scenario);
  }
  if (options.substrate.product?.trim()) {
    query.set("product", options.substrate.product);
  }
  if (options.substrate.agentHost?.trim()) {
    query.set("agentHost", options.substrate.agentHost);
  }
  if (options.substrate.licenseType?.trim()) {
    query.set("licenseType", options.substrate.licenseType);
  }
  if (typeof options.substrate.isEdu === "boolean") {
    query.set("isEdu", options.substrate.isEdu ? "true" : "false");
  }
  if (options.substrate.agent?.trim()) {
    query.set("agent", options.substrate.agent);
  }
  if (options.substrate.variants?.trim()) {
    query.set("variants", options.substrate.variants);
  }
  if (options.temporaryChat) {
    query.set("disableMemory", "1");
  }

  hubUri.search = query.toString();
  return hubUri;
}

export function buildInvocationPayload(
  request: ParsedOpenAiRequest,
  conversationId: string,
  sessionId: string,
  clientRequestId: string,
  isStartOfSession: boolean,
  options: WrapperOptions,
  observability: BridgeObservability | null = null,
  messageAnnotations: JsonObject[] = [],
): JsonObject {
  const capabilities = resolveSubstrateCapabilities(
    request.model,
    options.substrate,
  );
  const sendText = buildPromptWithAdditionalContext(request);
  const truncatedSendText = truncateSubstrateSendText(
    sendText,
    options.substrate.maxSendChars ?? 0,
    options.substrate.truncateBeforeSending ?? false,
  );
  if (truncatedSendText.length < sendText.length) {
    observability?.record("truncation", {
      originalChars: sendText.length,
      retainedChars: truncatedSendText.length,
      removedChars: sendText.length - truncatedSendText.length,
    });
  }
  const message: JsonObject = {
    author: "user",
    text: truncatedSendText,
    inputMethod: "Keyboard",
    messageType: capabilities.messageType,
    requestId: clientRequestId,
    messageId: randomUUID(),
    locale: capabilities.locale,
    experienceType: capabilities.experienceType,
  };

  if (capabilities.entityAnnotationTypes.length > 0) {
    message.entityAnnotationTypes = capabilities.entityAnnotationTypes;
  }

  const locationInfo = buildLocationInfo(request.locationHint);
  if (locationInfo) {
    message.locationInfo = locationInfo;
  }
  if (messageAnnotations.length > 0) {
    message.messageAnnotations = messageAnnotations;
  }

  const argument: JsonObject = {
    source: options.substrate.source?.trim() || "officeweb",
    tone: capabilities.tone,
    clientCorrelationId: clientRequestId,
    sessionId,
    conversationId,
    traceId: randomUUID().replaceAll("-", ""),
    isStartOfSession,
    streamingMode: capabilities.streamingMode,
    spokenTextMode: capabilities.spokenTextMode,
    options: {},
    extraExtensionParameters: {},
    sliceIds: [],
    threadLevelGptId: {},
    productThreadType: options.substrate.productThreadType?.trim() || "Office",
    clientInfo: {
      clientPlatform:
        options.substrate.clientPlatform?.trim() || "mcmcopilot-web",
      clientAppName: options.substrate.clientAppName?.trim() || "Office",
      clientEntrypoint:
        options.substrate.clientEntrypoint?.trim() ||
        "mcmcopilot-officeweb",
      clientSessionId: sessionId,
      ProductCategory: options.substrate.productCategory?.trim() || "Chat",
      clientAppType: options.substrate.clientAppType?.trim() || "Web",
      productEntryPoint:
        options.substrate.productEntryPoint?.trim() || "ChatPanel",
      deviceOS: options.substrate.deviceOS?.trim() || "Unknown",
      deviceType: options.substrate.deviceType?.trim() || "Desktop",
    },
    message,
    plugins: request.hostedWebSearch
      ? [{ Id: "BingWebSearch", Source: "BuiltIn" }]
      : [],
    isSbsSupported: true,
    renderReferencesBehindEOS: true,
    // Verified against the live substrate: `disconnectBehavior: "stop"` is
    // rejected and fails every invocation with a server-side Chat error, so
    // "continue" is the only accepted value. Cancellation therefore relies on
    // the CancelInvocation frame plus a hard socket close rather than on this
    // field. Do not change this without a live substrate turn to prove it.
    disconnectBehavior: "continue",
  };

  if (capabilities.optionsSets.length > 0) {
    argument.optionsSets = capabilities.optionsSets;
  }
  if (capabilities.allowedMessageTypes.length > 0) {
    argument.allowedMessageTypes = capabilities.allowedMessageTypes;
  }
  if (request.contextualResources) {
    argument.contextualResources = request.contextualResources;
  }

  return {
    arguments: [argument],
    invocationId: DefaultSubstrateInvocationId,
    target: capabilities.invocationTarget,
    type: capabilities.invocationType,
  };
}

export type SubstrateDeadlines = {
  invocationTimeoutMs: number;
  handshakeTimeoutMs: number;
  turnDeadlineMs: number;
};

export function resolveSubstrateDeadlines(
  substrate: SubstrateOptions,
  nowMs: number,
): SubstrateDeadlines {
  const invocationTimeoutMs =
    (substrate.invocationTimeoutSeconds > 0
      ? substrate.invocationTimeoutSeconds
      : 120) * 1000;
  const handshakeTimeoutMs =
    (substrate.handshakeTimeoutSeconds ?? invocationTimeoutMs / 1000) * 1000;
  const turnDeadlineMs =
    nowMs + (substrate.turnTimeoutSeconds ?? invocationTimeoutMs / 1000) * 1000;
  return { invocationTimeoutMs, handshakeTimeoutMs, turnDeadlineMs };
}

function buildPromptWithAdditionalContext(
  request: ParsedOpenAiRequest,
): string {
  if (request.additionalContext.length === 0) {
    return request.promptText;
  }
  const lines = ["Context:"];
  for (const ctx of request.additionalContext) {
    if (!ctx.text.trim()) {
      continue;
    }
    lines.push(`${ctx.description ? `${ctx.description}: ` : ""}${ctx.text}`);
  }
  lines.push("", `User: ${request.promptText}`);
  return lines.join("\n");
}

// Substrate silently drops or disconnects on oversized invocation payloads.
// When the combined prompt exceeds the configured ceiling we truncate the
// middle (context/history) and keep both ends, because the trailing text holds
// the user's actual prompt (`User: ...`) which must never be cut.
export function truncateSubstrateSendText(
  text: string,
  maxChars: number,
  enabled: boolean,
): string {
  if (!enabled || maxChars <= 0 || text.length <= maxChars) {
    return text;
  }
  const removed = text.length - maxChars;
  const marker = `\n\n…[${removed} characters truncated to fit Substrate limits]…\n\n`;
  if (marker.length >= maxChars) {
    throw new Error(
      "Substrate prompt limit is too small to preserve required context safely.",
    );
  }
  const userBoundary = text.lastIndexOf("\n\nUser: ");
  const requiredTail = userBoundary >= 0 ? text.slice(userBoundary) : text;
  const budget = maxChars - marker.length;
  if (requiredTail.length + marker.length >= maxChars) {
    if (userBoundary >= 0) {
      throw new Error(
        "Substrate prompt cannot be reduced without truncating the current user turn.",
      );
    }
    // Unstructured callers have no reliable user boundary. Preserve both the
    // request instructions and its tail instead of silently discarding either.
    const headChars = Math.floor(budget * 0.35);
    return (
      text.slice(0, headChars) +
      marker +
      text.slice(text.length - (budget - headChars))
    );
  }
  const tailChars = Math.max(requiredTail.length, Math.ceil(budget * 0.65));
  const headChars = budget - tailChars;
  return (
    text.slice(0, headChars) + marker + text.slice(text.length - tailChars)
  );
}

function buildLocationInfo(locationHint: JsonObject): JsonObject | null {
  const timeZone = tryGetString(locationHint, "timeZone");
  if (!timeZone) {
    return null;
  }
  const locationInfo: JsonObject = {
    timeZone,
    timeZoneOffset: resolveTimeZoneOffsetMinutes(timeZone),
  };
  const countryOrRegion = tryGetString(locationHint, "countryOrRegion");
  if (countryOrRegion) {
    locationInfo.countryOrRegion = countryOrRegion;
  }
  return locationInfo;
}

function resolveTimeZoneOffsetMinutes(timeZoneId: string): number {
  try {
    const now = new Date();
    const zoned = new Date(
      now.toLocaleString("en-US", { timeZone: timeZoneId }),
    );
    return Math.round((zoned.getTime() - now.getTime()) / 60000);
  } catch {
    return 0;
  }
}

const DefaultSubstrateInvocationId = "0";
// SignalR Hub Protocol CancelInvocation (type 5) uses the StreamInvocation ID:
// https://github.com/dotnet/aspnetcore/blob/main/src/SignalR/docs/specs/HubProtocol.md
const SignalRCancelInvocationType = 5;
const SubstrateCancelAckTimeoutMs = 500;

export function splitFrames(payload: string): string[] {
  return payload
    .split("\u001e")
    .map((frame) => frame.trim())
    .filter((frame) => frame.length > 0);
}

type SettledPromise<T> = {
  completed: boolean;
  value?: T;
  error?: unknown;
};

type SubstrateCancellationAcknowledgement = {
  wait: Promise<boolean>;
  dispose: () => void;
};

function settlePromiseWithin<T>(
  promise: Promise<T>,
  timeoutMs: number,
): Promise<SettledPromise<T>> {
  return new Promise<SettledPromise<T>>((resolve) => {
    let settled = false;
    const finish = (result: SettledPromise<T>) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      resolve(result);
    };
    const timer = setTimeout(
      () => finish({ completed: false }),
      timeoutMs,
    );
    promise.then(
      (value) => finish({ completed: true, value }),
      (error) => finish({ completed: true, error }),
    );
  });
}

function createSubstrateCancellationAcknowledgement(
  ws: WebSocket,
  invocationId: string,
  timeoutMs: number,
): SubstrateCancellationAcknowledgement {
  let settled = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let resolveWait!: (acknowledged: boolean) => void;
  let removeMessageListener = () => {};
  const wait = new Promise<boolean>((resolve) => {
    resolveWait = resolve;
  });
  const finish = (acknowledged: boolean) => {
    if (settled) {
      return;
    }
    settled = true;
    if (timer) {
      clearTimeout(timer);
      timer = null;
    }
    removeMessageListener();
    resolveWait(acknowledged);
  };
  try {
    removeMessageListener = listenToWebSocketEvent(ws, "message", (data) => {
      const payload = webSocketMessageToString(data);
      if (!payload) {
        return;
      }
      for (const frame of splitFrames(payload)) {
        const json = tryParseJsonObject(frame);
        if (
          tryGetInt(json, "type") === 3 &&
          tryGetString(json, "invocationId") === invocationId
        ) {
          finish(true);
          return;
        }
      }
    });
  } catch {
    // Test doubles and alternate WebSocket implementations may not expose events.
  }
  timer = setTimeout(() => finish(false), timeoutMs);
  return {
    wait,
    dispose: () => finish(false),
  };
}

function webSocketMessageToString(data: unknown): string {
  if (typeof data === "string") {
    return data;
  }
  if (Buffer.isBuffer(data)) {
    return data.toString("utf8");
  }
  if (Array.isArray(data)) {
    return Buffer.concat(data).toString("utf8");
  }
  if (data instanceof ArrayBuffer) {
    return Buffer.from(data).toString("utf8");
  }
  if (ArrayBuffer.isView(data)) {
    return Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString(
      "utf8",
    );
  }
  return String(data ?? "");
}

function parseRawJsonValue(rawBody: string): JsonValue {
  try {
    return JSON.parse(rawBody) as JsonValue;
  } catch {
    return { rawText: rawBody };
  }
}

function buildSubstrateTranscriptPayload(transcript: string[]): JsonValue {
  const frames = transcript.flatMap((payload) =>
    splitFrames(payload).map((frame) => {
      const parsed = tryParseJsonObject(frame);
      return parsed ?? { rawText: frame };
    }),
  );

  return {
    streamType: "signalr",
    frameCount: frames.length,
    frames,
  };
}

async function connectWithAbort(
  connectPromise: Promise<WebSocket>,
  signal: AbortSignal | null,
): Promise<WebSocket | Error> {
  if (!signal) {
    return connectPromise.catch((error) => error as Error);
  }
  if (signal.aborted) {
    const error = new Error("WebSocket connection aborted.");
    error.name = "AbortError";
    return error;
  }

  return new Promise<WebSocket | Error>((resolve) => {
    let settled = false;
    const finish = (value: WebSocket | Error) => {
      if (settled) {
        if (!(value instanceof Error)) {
          try {
            if (
              typeof (value as { terminate?: () => void }).terminate ===
              "function"
            ) {
              (value as { terminate: () => void }).terminate();
            } else {
              value.close(1000, "client disconnected");
            }
          } catch {
            // ignore
          }
        }
        return;
      }
      settled = true;
      signal.removeEventListener("abort", onAbort);
      resolve(value);
    };
    const onAbort = () => {
      const error = new Error("WebSocket connection aborted.");
      error.name = "AbortError";
      finish(error);
    };
    signal.addEventListener("abort", onAbort, { once: true });
    connectPromise.then(finish).catch((error) => finish(error as Error));
  });
}

export async function connectWebSocket(
  url: URL,
  origin: string | undefined,
  timeoutMs: number,
  keepAliveMs: number,
): Promise<WebSocket> {
  return new Promise<WebSocket>((resolve, reject) => {
    const ws = new WebSocket(url, {
      handshakeTimeout: timeoutMs,
      headers: origin ? { Origin: origin } : undefined,
    });
    installWebSocketErrorSink(ws);
    const timeout = setTimeout(() => {
      try {
        if (typeof (ws as { terminate?: () => void }).terminate === "function") {
          ws.terminate();
        } else {
          ws.close();
        }
      } catch {
        // ignore
      }
      reject(new Error("timeout"));
    }, timeoutMs);

    listenToWebSocketEvent(ws, "open", () => {
      clearTimeout(timeout);
      if (keepAliveMs > 0) {
        const timer = setInterval(() => {
          if (ws.readyState !== WebSocket.OPEN) {
            clearInterval(timer);
            return;
          }
          try {
            if (typeof (ws as { ping?: () => void }).ping === "function") {
              ws.ping();
            }
          } catch {
            clearInterval(timer);
          }
        }, keepAliveMs);
        listenToWebSocketEvent(ws, "close", () => clearInterval(timer), true);
      }
      resolve(ws);
    }, true);
    listenToWebSocketEvent(ws, "error", (error) => {
      clearTimeout(timeout);
      reject(normalizeWebSocketError(error));
    }, true);
  });
}

export async function sendFrame(
  ws: WebSocket,
  requestUri: URL,
  logger: DebugMarkdownLogger,
  frame: JsonObject,
): Promise<void> {
  const payload = `${JSON.stringify(frame)}\u001e`;
  await logger.logSubstrateFrame(requestUri.toString(), "request", payload);
  await new Promise<void>((resolve, reject) => {
    try {
      if (ws.send.length >= 2) {
        ws.send(payload, (error) => {
          if (error) {
            reject(error);
          } else {
            resolve();
          }
        });
        return;
      }
      ws.send(payload);
      resolve();
    } catch (error) {
      reject(error);
    }
  });
}

export function createWebSocketReceiver(ws: WebSocket): {
  next: (timeoutMs: number) => Promise<string | null>;
  dispose: () => void;
} {
  const queue: Array<string | null> = [];
  const waiters: Array<(value: string | null) => void> = [];
  let disposed = false;

  const flush = (value: string | null) => {
    if (waiters.length > 0) {
      const waiter = waiters.shift();
      if (waiter) {
        waiter(value);
      }
      return;
    }
    queue.push(value);
  };

  const onMessage = (data: unknown) => {
    if (disposed) {
      return;
    }
    if (typeof data === "string") {
      flush(data);
      return;
    }
    if (Buffer.isBuffer(data)) {
      flush(data.toString("utf8"));
      return;
    }
    if (Array.isArray(data)) {
      flush(Buffer.concat(data).toString("utf8"));
      return;
    }
    if (data instanceof ArrayBuffer) {
      flush(Buffer.from(data).toString("utf8"));
      return;
    }
    if (ArrayBuffer.isView(data)) {
      flush(
        Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString("utf8"),
      );
      return;
    }
    flush(String(data ?? ""));
  };

  const onClose = () => {
    if (disposed) {
      return;
    }
    flush(null);
  };

  const onError = () => {
    if (disposed) {
      return;
    }
    flush(null);
  };

  const offMessage = listenToWebSocketEvent(ws, "message", onMessage);
  const offClose = listenToWebSocketEvent(ws, "close", onClose);
  const offError = listenToWebSocketEvent(ws, "error", onError);

  return {
    next: (timeoutMs: number) => {
      if (queue.length > 0) {
        return Promise.resolve(queue.shift() ?? null);
      }
      if (disposed) {
        return Promise.resolve(null);
      }
      return new Promise<string | null>((resolve) => {
        const waiter = (value: string | null) => {
          clearTimeout(timer);
          const index = waiters.indexOf(waiter);
          if (index >= 0) {
            waiters.splice(index, 1);
          }
          resolve(value);
        };
        const timer = setTimeout(() => {
          const index = waiters.indexOf(waiter);
          if (index >= 0) {
            waiters.splice(index, 1);
          }
          resolve(null);
        }, timeoutMs);
        waiters.push(waiter);
      });
    },
    dispose: () => {
      if (disposed) {
        return;
      }
      disposed = true;
      offMessage();
      offClose();
      offError();
      while (waiters.length > 0) {
        const waiter = waiters.shift();
        waiter?.(null);
      }
      queue.length = 0;
    },
  };
}

function listenToWebSocketEvent(
  ws: WebSocket,
  event: "open" | "message" | "close" | "error",
  handler: (...args: unknown[]) => void,
  once = false,
): () => void {
  const emitter = ws as WebSocket & {
    on?: (event: string, handler: (...args: unknown[]) => void) => void;
    once?: (event: string, handler: (...args: unknown[]) => void) => void;
    off?: (event: string, handler: (...args: unknown[]) => void) => void;
    addListener?: (event: string, handler: (...args: unknown[]) => void) => void;
    removeListener?: (event: string, handler: (...args: unknown[]) => void) => void;
    addEventListener?: (
      event: string,
      handler: (event: { data?: unknown; error?: unknown }) => void,
      options?: { once?: boolean },
    ) => void;
    removeEventListener?: (
      event: string,
      handler: (event: { data?: unknown; error?: unknown }) => void,
    ) => void;
    onmessage?: ((event: { data?: unknown }) => void) | null;
    onerror?: ((event: { error?: unknown }) => void) | null;
  };

  if (once && typeof emitter.once === "function") {
    emitter.once(event, handler);
    return () => {
      if (typeof emitter.off === "function") {
        emitter.off(event, handler);
      } else {
        emitter.removeListener?.(event, handler);
      }
    };
  }

  if (typeof emitter.on === "function") {
    emitter.on(event, handler);
    return () => {
      if (typeof emitter.off === "function") {
        emitter.off(event, handler);
      } else {
        emitter.removeListener?.(event, handler);
      }
    };
  }

  if (typeof emitter.addListener === "function") {
    emitter.addListener(event, handler);
    return () => emitter.removeListener?.(event, handler);
  }

  if (typeof emitter.addEventListener === "function") {
    const wrapped = (domEvent: { data?: unknown; error?: unknown }) => {
      if (event === "message") {
        handler(domEvent.data);
        return;
      }
      if (event === "error") {
        handler(domEvent.error ?? domEvent);
        return;
      }
      handler(domEvent);
    };
    emitter.addEventListener(event, wrapped, once ? { once: true } : undefined);
    return () => emitter.removeEventListener?.(event, wrapped);
  }

  // Bun's native WebSocket can expose property handlers without the Node or
  // DOM listener methods. Support that shape for the live Substrate client.
  if (event === "message" && "onmessage" in emitter) {
    const previous = emitter.onmessage;
    const wrapped = (domEvent: { data?: unknown }) => handler(domEvent.data);
    emitter.onmessage = wrapped;
    return () => {
      if (emitter.onmessage === wrapped) emitter.onmessage = previous ?? null;
    };
  }
  if (event === "error" && "onerror" in emitter) {
    const previous = emitter.onerror;
    const wrapped = (domEvent: { error?: unknown }) => handler(domEvent.error ?? domEvent);
    emitter.onerror = wrapped;
    return () => {
      if (emitter.onerror === wrapped) emitter.onerror = previous ?? null;
    };
  }

  throw new Error(`Unsupported WebSocket event API for '${event}'.`);
}

function normalizeWebSocketError(error: unknown): Error {
  if (error instanceof Error) {
    return error;
  }
  if (
    typeof error === "object" &&
    error !== null &&
    "message" in error &&
    typeof error.message === "string" &&
    error.message.trim()
  ) {
    return new Error(error.message);
  }
  if (
    typeof error === "object" &&
    error !== null &&
    "type" in error &&
    typeof error.type === "string" &&
    error.type.trim()
  ) {
    return new Error(`WebSocket error event: ${error.type}`);
  }
  return new Error(String(error));
}

function installWebSocketErrorSink(ws: WebSocket): void {
  try {
    listenToWebSocketEvent(ws, "error", () => {
      // Keep an error listener attached so failed websocket handshakes return
      // normal request errors instead of crashing the Bun process.
    });
  } catch {
    // Best effort only; runtimes without a compatible listener API still rely
    // on the explicit connect/receiver error handling paths.
  }
}

function extractSubstrateAssistantText(envelope: JsonObject): string | null {
  const messages = collectMessageObjects(envelope);
  let fallback: string | null = null;
  for (const message of messages) {
    if ((tryGetString(message, "author") ?? "").toLowerCase() !== "bot") {
      continue;
    }
    const messageType = (
      tryGetString(message, "messageType") ?? "Chat"
    ).toLowerCase();
    if (messageType !== "chat") {
      continue;
    }

    const text =
      tryGetString(message, "text") ??
      tryGetString(message, "hiddenText") ??
      tryGetString(message, "spokenText");
    if (text) {
      fallback = text;
    }
  }
  return fallback ?? extractSubstrateResultMessage(envelope);
}

function extractSubstrateDisengagedMessage(envelope: JsonObject): string | null {
  for (const message of collectMessageObjects(envelope)) {
    if ((tryGetString(message, "author") ?? "").toLowerCase() !== "bot") {
      continue;
    }
    if (
      (tryGetString(message, "messageType") ?? "").toLowerCase() !== "disengaged"
    ) {
      continue;
    }
    return (
      tryGetString(message, "text") ??
      tryGetString(message, "hiddenText") ??
      tryGetString(message, "spokenText") ??
      "M365 Copilot disengaged from the request."
    );
  }
  return null;
}

function extractSubstrateDeltaText(envelope: JsonObject): string | null {
  const args = envelope.arguments;
  if (!Array.isArray(args)) {
    return null;
  }
  for (const arg of args) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const delta = tryGetString(arg, "writeAtCursor");
    if (delta) {
      return delta;
    }
  }
  return null;
}

function extractSubstrateConversationId(envelope: JsonObject): string | null {
  const direct = tryGetString(envelope, "conversationId");
  if (direct) {
    return direct;
  }

  const item = envelope.item;
  if (isJsonObject(item)) {
    const itemId = tryGetString(item, "conversationId");
    if (itemId) {
      return itemId;
    }
  }

  const args = envelope.arguments;
  if (!Array.isArray(args)) {
    return null;
  }
  for (const arg of args) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const argId = tryGetString(arg, "conversationId");
    if (argId) {
      return argId;
    }
    const argItem = arg.item;
    if (isJsonObject(argItem)) {
      const argItemId = tryGetString(argItem, "conversationId");
      if (argItemId) {
        return argItemId;
      }
    }
  }
  return null;
}

function extractSubstrateRequestId(envelope: JsonObject): string | null {
  const direct = tryGetString(envelope, "requestId");
  if (direct) {
    return direct;
  }

  const item = envelope.item;
  if (isJsonObject(item)) {
    const itemRequestId = tryGetString(item, "requestId");
    if (itemRequestId) {
      return itemRequestId;
    }
  }

  const args = envelope.arguments;
  if (!Array.isArray(args)) {
    return null;
  }
  for (const arg of args) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const argRequestId = tryGetString(arg, "requestId");
    if (argRequestId) {
      return argRequestId;
    }
    const argItem = arg.item;
    if (isJsonObject(argItem)) {
      const argItemRequestId = tryGetString(argItem, "requestId");
      if (argItemRequestId) {
        return argItemRequestId;
      }
    }
  }

  return null;
}

function buildSubstrateFrameIdentity(
  envelope: JsonObject,
  fallback: string,
): string {
  const requestId = extractSubstrateRequestId(envelope) ?? "";
  const messageId = extractSubstrateBotMessageId(envelope) ?? "";
  const frameType = tryGetInt(envelope, "type") ?? -1;
  if (requestId || messageId) {
    return [frameType, requestId, messageId, fallback].join(":");
  }
  return fallback;
}

function extractSubstrateBotMessageId(envelope: JsonObject): string | null {
  const messages = collectMessageObjects(envelope);
  let lastMessageId: string | null = null;
  for (const message of messages) {
    const author = (tryGetString(message, "author") ?? "").toLowerCase();
    if (author !== "bot") {
      continue;
    }
    const messageId = tryGetString(message, "messageId");
    if (messageId) {
      lastMessageId = messageId;
    }
  }
  return lastMessageId;
}

function extractSubstrateResultMessage(envelope: JsonObject): string | null {
  const item = envelope.item;
  if (isJsonObject(item) && isJsonObject(item.result)) {
    const itemMessage = tryGetString(item.result, "message");
    if (itemMessage) {
      return itemMessage;
    }
  }
  if (isJsonObject(envelope.result)) {
    return tryGetString(envelope.result, "message");
  }
  return null;
}

function extractSubstrateResultValue(envelope: JsonObject): string | null {
  const item = envelope.item;
  if (isJsonObject(item) && isJsonObject(item.result)) {
    const itemValue = tryGetString(item.result, "value");
    if (itemValue) {
      return itemValue;
    }
  }
  if (isJsonObject(envelope.result)) {
    return tryGetString(envelope.result, "value");
  }
  return null;
}

function hasSubstrateResultValue(envelope: JsonObject): boolean {
  if (Object.prototype.hasOwnProperty.call(envelope, "result")) {
    return true;
  }
  return (
    isJsonObject(envelope.item) &&
    Object.prototype.hasOwnProperty.call(envelope.item, "result")
  );
}

function isSubstrateResultSuccess(resultValue: string): boolean {
  const normalized = resultValue.toLowerCase();
  return normalized === "success" || normalized === "apologyresponsereturned";
}

function hasCompleteSimulatedPayload(assistantText: string): boolean {
  const chatPayload = tryExtractSimulatedResponsePayload(
    assistantText,
    "chat.completions",
  );
  if (chatPayload) {
    const assistantResponse =
      tryBuildAssistantResponseFromChatCompletionPayload(chatPayload);
    if (!assistantResponse) {
      return false;
    }
    // Tool-call payloads tend to continue changing across subsequent frames.
    // Early completion is safer for plain assistant text responses only.
    if (assistantResponse.toolCalls.length > 0) {
      return false;
    }
    return Boolean(assistantResponse.content?.trim());
  }

  const responsesPayload = tryExtractSimulatedResponsePayload(
    assistantText,
    "responses",
  );
  if (!responsesPayload) {
    return false;
  }

  const output = responsesPayload.output;
  if (!Array.isArray(output) || output.length === 0) {
    return false;
  }

  let hasMessageText = false;
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type === "function_call") {
      return false;
    }
    if (type === "message") {
      const content = item.content;
      if (Array.isArray(content)) {
        for (const part of content) {
          if (!isJsonObject(part)) {
            continue;
          }
          const text =
            tryGetString(part, "text") ?? tryGetString(part, "output_text");
          if (text?.trim()) {
            hasMessageText = true;
            break;
          }
        }
      }
    }
  }

  if (hasMessageText) {
    return true;
  }

  return Boolean(tryGetString(responsesPayload, "output_text")?.trim());
}

function resolveCompleteSimulatedAssistantText(
  latestAssistantText: string | null,
  accumulatedAssistantText: string,
): string | null {
  if (latestAssistantText && hasCompleteSimulatedPayload(latestAssistantText)) {
    return latestAssistantText;
  }

  if (
    accumulatedAssistantText &&
    hasCompleteSimulatedPayload(accumulatedAssistantText)
  ) {
    return accumulatedAssistantText;
  }

  return null;
}

function buildNormalizedConversation(
  conversationId: string,
  prompt: string,
  assistantText: string,
): JsonObject {
  return {
    id: conversationId,
    messages: [
      { author: "user", text: prompt },
      { author: "assistant", text: assistantText },
    ],
  };
}

function collectMessageObjects(envelope: JsonObject): JsonObject[] {
  const messages: JsonObject[] = [];
  const pushArray = (value: JsonValue | undefined) => {
    if (!Array.isArray(value)) {
      return;
    }
    for (const item of value) {
      if (isJsonObject(item)) {
        messages.push(item);
      }
    }
  };

  pushArray(envelope.messages);
  if (isJsonObject(envelope.item)) {
    pushArray(envelope.item.messages);
  }
  const args = envelope.arguments;
  if (Array.isArray(args)) {
    for (const arg of args) {
      if (!isJsonObject(arg)) {
        continue;
      }
      pushArray(arg.messages);
      if (isJsonObject(arg.item)) {
        pushArray(arg.item.messages);
      }
    }
  }
  return messages;
}

export function summarizeUpstreamFailure(
  statusCode: number,
  responseBody: string | null,
  fallbackMessage: string,
): { statusCode: number; message: string } {
  const details = extractGraphErrorMessage(responseBody);
  const message = details ? `${fallbackMessage} ${details}` : fallbackMessage;
  const normalizedStatusCode =
    statusCode >= 400 && statusCode <= 599 ? statusCode : 502;
  return { statusCode: normalizedStatusCode, message };
}
