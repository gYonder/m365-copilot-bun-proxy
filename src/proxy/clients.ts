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

export class CopilotSubstrateClient {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
    private readonly sessionStore = new SubstrateSessionStore(
      options.conversationTtlMinutes,
    ),
    private readonly connect: SubstrateWebSocketConnector = connectWebSocket,
    private readonly createReceiver: SubstrateReceiverFactory = createWebSocketReceiver,
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
    try {
      return await this.sessionStore.runExclusive(
        sessionId,
        () =>
          this.chatCoreUnlocked(
            authorizationHeader,
            conversationId,
            sessionId,
            request,
            isStartOfSession,
            onStreamUpdate,
            signal,
            taskDeadlineMs,
          ),
        signal,
        taskDeadlineMs,
      );
    } catch (error) {
      if (error instanceof TurnQueueWaitError) {
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

    const clientRequestId = randomUUID();
    const requestUri = buildSubstrateHubUri(
      this.options,
      objectId,
      tenantId,
      rawToken,
      clientRequestId,
      sessionId,
      conversationId,
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
    let resolvedConversationId = conversationId;
    const receiver = this.createReceiver(ws);
    const onAbort = () => {
      try {
        ws.close(1000, "client disconnected");
      } catch {
        // ignore
      }
    };
    if (signal) {
      if (signal.aborted) {
        onAbort();
      } else {
        signal.addEventListener("abort", onAbort, { once: true });
      }
    }
    try {
      await sendFrame(ws, requestUri, this.logger, {
        protocol: "json",
        version: 1,
      });
      const handshakeDeadlineMs = Date.now() + remainingHandshakeMs;
      const handshakePayload = await receiver.next(remainingHandshakeMs);
      if (handshakePayload === null) {
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

      await sendFrame(ws, requestUri, this.logger, { type: 6 });
      invocationPayload = buildInvocationPayload(
        request,
        conversationId,
        sessionId,
        clientRequestId,
        isStartOfSession,
        this.options,
      );
      if (onStreamUpdate) {
        await onStreamUpdate({
          deltaText: null,
          conversationId: resolvedConversationId,
          upstreamRequestPayload: invocationPayload,
          upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
        });
      }
      await sendFrame(
        ws,
        requestUri,
        this.logger,
        invocationPayload,
      );

      let assistantText = "";
      let deltaBuilder = "";
      let responseError: string | null = null;
      let responseErrorCode: string | null = null;
      let completed = false;
      let activeBotMessageId: string | null = null;
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
        if (signal?.aborted) {
          break;
        }
        const remainingMs = turnDeadlineMs - Date.now();
        if (remainingMs <= 0) {
          return buildFailure(504, "Substrate websocket turn timed out.");
        }
        const payload = await receiver.next(remainingMs);
        if (payload === null) {
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
        await this.logger.logSubstrateFrame(
          requestUri.toString(),
          "response",
          payload,
        );
        transcript.push(payload);
        if (onStreamUpdate) {
          await onStreamUpdate({
            deltaText: null,
            conversationId: resolvedConversationId,
            upstreamResponsePayload: buildSubstrateTranscriptPayload(transcript),
          });
        }

        for (const frame of splitFrames(payload)) {
          if (!frame.trim()) {
            continue;
          }
          const json = tryParseJsonObject(frame);
          if (!json) {
            continue;
          }
          const frameType = tryGetInt(json, "type");
          if (frameType === 6) {
            await sendFrame(ws, requestUri, this.logger, { type: 6 });
            continue;
          }

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
          }

          const deltaText = extractSubstrateDeltaText(json);
          if (deltaText) {
            deltaBuilder += deltaText;
            await logDeltaSource("writeAtCursor", deltaText);
            if (onStreamUpdate) {
              await onStreamUpdate({
                deltaText,
                conversationId: resolvedConversationId,
              });
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
              if (onStreamUpdate) {
                await onStreamUpdate({
                  deltaText: snapshotDelta,
                  conversationId: resolvedConversationId,
                });
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

          if (
            frameType !== null &&
            (frameType === 3 || frameType === 7)
          ) {
            completed = true;
            break;
          }
        }
      }

      try {
        ws.close(1000, "completed");
      } catch {
        // ignore
      }

      if (signal?.aborted) {
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
          responseErrorCode,
        );
      }

      if (!assistantText && deltaBuilder) {
        assistantText = deltaBuilder;
      }

      if (!assistantText) {
        return buildFailure(
          502,
          "Substrate chat returned no assistant content.",
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
      receiver.dispose();
      try {
        ws.close();
      } catch {
        // ignore
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
): JsonObject {
  const message: JsonObject = {
    author: "user",
    text: buildPromptWithAdditionalContext(request),
    inputMethod: "Keyboard",
    messageType: "Chat",
    requestId: clientRequestId,
    messageId: randomUUID(),
    locale: options.substrate.locale?.trim() || "en-US",
    experienceType: options.substrate.experienceType?.trim() || "Default",
  };

  if (options.substrate.entityAnnotationTypes.length > 0) {
    message.entityAnnotationTypes = options.substrate.entityAnnotationTypes
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }

  const locationInfo = buildLocationInfo(request.locationHint);
  if (locationInfo) {
    message.locationInfo = locationInfo;
  }

  const argument: JsonObject = {
    source: options.substrate.source?.trim() || "officeweb",
    tone: resolveSubstrateTone(request.model),
    clientCorrelationId: clientRequestId,
    sessionId,
    conversationId,
    traceId: randomUUID().replaceAll("-", ""),
    isStartOfSession,
    streamingMode: "ConciseWithPadding",
    spokenTextMode: "None",
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
    plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }],
    isSbsSupported: true,
    renderReferencesBehindEOS: true,
    disconnectBehavior: "continue",
  };

  if (options.substrate.optionsSets.length > 0) {
    argument.optionsSets = options.substrate.optionsSets
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }
  if (options.substrate.allowedMessageTypes.length > 0) {
    argument.allowedMessageTypes = options.substrate.allowedMessageTypes
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }
  if (request.contextualResources) {
    argument.contextualResources = request.contextualResources;
  }

  return {
    arguments: [argument],
    invocationId: "0",
    target: options.substrate.invocationTarget?.trim() || "update",
    type:
      options.substrate.invocationType > 0
        ? options.substrate.invocationType
        : 1,
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

export function resolveSubstrateTone(model: string | null | undefined): string {
  const normalizedModel = model?.trim().toLowerCase() ?? "";
  switch (normalizedModel) {
    case "gpt-5.6-sol":
      return "Gpt_5_6_Reasoning";
    default:
      return "magic";
  }
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

function splitFrames(payload: string): string[] {
  return payload
    .split("\u001e")
    .map((frame) => frame.trim())
    .filter((frame) => frame.length > 0);
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
            value.close(1000, "client disconnected");
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

async function connectWebSocket(
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

async function sendFrame(
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

function createWebSocketReceiver(ws: WebSocket): {
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
