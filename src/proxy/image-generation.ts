import { ConcurrencyLimiter } from "./concurrency";
import { randomUUID } from "node:crypto";
import {
  buildSubstrateHubUri,
  connectWebSocket,
  createWebSocketReceiver,
  sendFrame,
  splitFrames,
} from "./clients";
import { DebugMarkdownLogger } from "./logger";
import {
  extractBearerToken,
  isJsonObject,
  tryGetString,
  tryParseJsonObject,
  tryReadJwtPayload,
} from "./utils";
import type {
  ImageGenerationOptions,
  JsonObject,
  JsonValue,
  WrapperOptions,
} from "./types";

export type ImageArtifact = {
  url: string;
  revisedPrompt?: string | null;
};

export type ImageGenerationTransport = {
  generate(
    authorizationHeader: string,
    prompt: string,
    count: number,
    signal: AbortSignal,
  ): Promise<ImageArtifact[]>;
};

export type ImageTokenProvider = {
  resolveDesignerAuthorizationHeader(): Promise<string | null>;
};

export type ImageFetch = typeof fetch;

export class DesignerSubstrateImageTransport implements ImageGenerationTransport {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
  ) {}

  async generate(
    authorizationHeader: string,
    prompt: string,
    count: number,
    signal: AbortSignal,
  ): Promise<ImageArtifact[]> {
    const token = extractBearerToken(authorizationHeader);
    const claims = token ? tryReadJwtPayload(token) : null;
    const objectId = claims ? tryGetString(claims, "oid") : null;
    const tenantId = claims ? tryGetString(claims, "tid") : null;
    if (!token || !objectId || !tenantId) {
      throw new ImageGenerationError(
        401,
        "invalid_designer_token",
        "Designer authorization lacks required identity claims.",
      );
    }

    const conversationId = randomUUID();
    const sessionId = randomUUID();
    const requestId = randomUUID();
    const requestUri = buildSubstrateHubUri(
      this.options,
      objectId,
      tenantId,
      token,
      requestId,
      sessionId,
      conversationId,
    );
    const timeoutMs = this.options.imageGeneration.timeoutMs;
    const socket = await connectWebSocket(
      requestUri,
      this.options.substrate.origin ?? undefined,
      timeoutMs,
      this.options.substrate.keepAliveSeconds * 1_000,
    );
    const receiver = createWebSocketReceiver(socket);
    const closeForAbort = () => {
      try {
        socket.close(1000, "image generation cancelled");
      } catch {
        // Best effort cancellation.
      }
    };
    signal.addEventListener("abort", closeForAbort, { once: true });

    try {
      await sendFrame(socket, requestUri, this.logger, {
        protocol: "json",
        version: 1,
      });
      const handshake = await receiver.next(timeoutMs);
      if (!handshake) {
        throw new ImageGenerationError(
          504,
          "image_generation_timeout",
          "M365 image generation handshake timed out.",
        );
      }

      await sendFrame(socket, requestUri, this.logger, { type: 6 });
      await sendFrame(socket, requestUri, this.logger, {
        arguments: [
          {
            source: this.options.substrate.source,
            sessionId,
            conversationId,
            clientCorrelationId: requestId,
            message: {
              author: "user",
              text: prompt,
              messageType: "GenerateGraphicArt",
              requestId,
              messageId: randomUUID(),
              locale: this.options.substrate.locale,
            },
            imageGeneration: { count },
            allowedMessageTypes: ["Progress", "GraphicArt", "EndOfRequest"],
          },
        ],
        invocationId: "0",
        target: this.options.substrate.invocationTarget,
        type: this.options.substrate.invocationType,
      });

      const artifacts: ImageArtifact[] = [];
      const seen = new Set<string>();
      const deadline = Date.now() + timeoutMs;
      let completed = false;
      while (!completed && Date.now() < deadline) {
        if (signal.aborted) throw new DOMException("aborted", "AbortError");
        const payload = await receiver.next(Math.max(1, deadline - Date.now()));
        if (!payload) break;
        for (const frame of splitFrames(payload)) {
          const parsed = tryParseJsonObject(frame);
          if (!parsed) continue;
          const frameType = parsed.type;
          if (frameType === 6) {
            await sendFrame(socket, requestUri, this.logger, { type: 6 });
            continue;
          }
          collectGraphicArtArtifacts(parsed, artifacts, seen);
          if (frameType === 7) {
            throw new ImageGenerationError(
              502,
              "image_generation_failed",
              "M365 image generation ended with an error.",
            );
          }
          if (frameType === 3) completed = true;
        }
      }

      if (!completed) {
        throw new ImageGenerationError(
          504,
          "image_generation_timeout",
          "M365 image generation did not reach a terminal state.",
        );
      }
      return artifacts.slice(0, count);
    } finally {
      signal.removeEventListener("abort", closeForAbort);
      receiver.dispose();
      try {
        socket.close(1000, "completed");
      } catch {
        // Socket may already be closed.
      }
    }
  }
}

function collectGraphicArtArtifacts(
  value: JsonValue,
  artifacts: ImageArtifact[],
  seen: Set<string>,
): void {
  if (Array.isArray(value)) {
    for (const item of value) collectGraphicArtArtifacts(item, artifacts, seen);
    return;
  }
  if (!isJsonObject(value)) return;

  const messageType = (tryGetString(value, "messageType") ?? "").toLowerCase();
  if (messageType === "graphicart" || "artifactUrl" in value || "imageUrl" in value) {
    const url =
      tryGetString(value, "artifactUrl") ??
      tryGetString(value, "imageUrl") ??
      tryGetString(value, "url");
    if (url && !seen.has(url)) {
      seen.add(url);
      artifacts.push({
        url,
        revisedPrompt:
          tryGetString(value, "revisedPrompt") ??
          tryGetString(value, "revised_prompt"),
      });
    }
  }

  for (const nested of Object.values(value)) {
    collectGraphicArtArtifacts(nested, artifacts, seen);
  }
}

export class ImageGenerationError extends Error {
  constructor(
    readonly statusCode: number,
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = "ImageGenerationError";
  }
}

export class ImageGenerationService {
  private readonly limiter: ConcurrencyLimiter;

  constructor(
    private readonly options: ImageGenerationOptions,
    private readonly tokenProvider: ImageTokenProvider,
    private readonly transport: ImageGenerationTransport,
    private readonly fetchArtifact: ImageFetch = fetch,
  ) {
    this.limiter = new ConcurrencyLimiter(options.concurrencyLimit, options.timeoutMs);
  }

  async create(request: Request): Promise<Response> {
    if (!this.options.enabled) {
      return errorResponse(404, "image_generation_disabled", "Image generation is disabled.");
    }

    let body: JsonObject;
    try {
      const candidate = (await request.json()) as JsonValue;
      if (!isJsonObject(candidate)) throw new Error("invalid body");
      body = candidate;
    } catch {
      return errorResponse(400, "invalid_request_error", "The request body must be a JSON object.");
    }

    const prompt = tryGetString(body, "prompt")?.trim() ?? "";
    if (!prompt) {
      return errorResponse(400, "missing_prompt", "A non-empty prompt is required.");
    }
    if (prompt.length > this.options.maxPromptChars) {
      return errorResponse(400, "prompt_too_long", "The image prompt exceeds the configured limit.");
    }

    const rawCount = body.n ?? 1;
    const count = typeof rawCount === "number" && Number.isInteger(rawCount) ? rawCount : 0;
    if (count < 1 || count > this.options.maxImages) {
      return errorResponse(400, "invalid_image_count", "The requested image count is outside the configured limit.");
    }

    const responseFormat = tryGetString(body, "response_format") ?? "b64_json";
    if (responseFormat !== "b64_json") {
      return errorResponse(400, "unsupported_response_format", "Only b64_json image responses are supported.");
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.options.timeoutMs);
    const onAbort = () => controller.abort();
    request.signal.addEventListener("abort", onAbort, { once: true });

    let release: (() => void) | null = null;
    try {
      release = await this.limiter.acquire(controller.signal, Date.now() + this.options.timeoutMs);
      const authorizationHeader = await this.tokenProvider.resolveDesignerAuthorizationHeader();
      if (!authorizationHeader) {
        throw new ImageGenerationError(401, "designer_auth_unavailable", "Designer authorization is unavailable.");
      }

      const artifacts = await this.transport.generate(
        authorizationHeader,
        prompt,
        count,
        controller.signal,
      );
      if (artifacts.length !== count) {
        throw new ImageGenerationError(502, "incomplete_image_result", "M365 returned an incomplete image result.");
      }

      const data = [];
      for (const artifact of artifacts) {
        const retrieved = await retrieveArtifact(
          artifact.url,
          this.options,
          this.fetchArtifact,
          controller.signal,
        );
        data.push({
          b64_json: Buffer.from(retrieved.bytes).toString("base64"),
          revised_prompt: artifact.revisedPrompt ?? undefined,
        });
      }

      return Response.json({ created: Math.floor(Date.now() / 1000), data });
    } catch (error) {
      if (error instanceof ImageGenerationError) {
        return errorResponse(error.statusCode, error.code, error.message);
      }
      if (controller.signal.aborted) {
        const cancelled = request.signal.aborted;
        return errorResponse(
          cancelled ? 499 : 504,
          cancelled ? "image_generation_cancelled" : "image_generation_timeout",
          cancelled ? "Image generation was cancelled." : "Image generation timed out.",
        );
      }
      return errorResponse(502, "image_generation_failed", "M365 image generation failed.");
    } finally {
      release?.();
      clearTimeout(timeout);
      request.signal.removeEventListener("abort", onAbort);
    }
  }
}

async function retrieveArtifact(
  rawUrl: string,
  options: ImageGenerationOptions,
  fetchArtifact: ImageFetch,
  signal: AbortSignal,
): Promise<{ bytes: Uint8Array; mimeType: string }> {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    throw new ImageGenerationError(502, "invalid_artifact_url", "M365 returned an invalid artifact URL.");
  }
  if (url.protocol !== "https:") {
    throw new ImageGenerationError(502, "invalid_artifact_url", "M365 returned an insecure artifact URL.");
  }

  const response = await fetchArtifact(url, { signal, redirect: "manual" });
  if (response.status >= 300 && response.status < 400) {
    throw new ImageGenerationError(502, "artifact_redirect_rejected", "Image artifact redirects are not allowed.");
  }
  if (!response.ok) {
    throw new ImageGenerationError(502, "artifact_retrieval_failed", "The generated image artifact could not be retrieved.");
  }

  const mimeType = (response.headers.get("content-type") ?? "").split(";", 1)[0].trim().toLowerCase();
  if (!options.allowedMimeTypes.includes(mimeType)) {
    throw new ImageGenerationError(502, "unexpected_artifact_type", "The generated artifact has an unsupported MIME type.");
  }

  const declaredLength = Number(response.headers.get("content-length") ?? "0");
  if (Number.isFinite(declaredLength) && declaredLength > options.maxArtifactBytes) {
    throw new ImageGenerationError(502, "artifact_too_large", "The generated image artifact exceeds the configured size limit.");
  }

  const reader = response.body?.getReader();
  if (!reader) {
    throw new ImageGenerationError(502, "artifact_retrieval_failed", "The generated image artifact has no response body.");
  }
  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const next = await reader.read();
    if (next.done) break;
    total += next.value.byteLength;
    if (total > options.maxArtifactBytes) {
      await reader.cancel().catch(() => {});
      throw new ImageGenerationError(502, "artifact_too_large", "The generated image artifact exceeds the configured size limit.");
    }
    chunks.push(next.value);
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return { bytes, mimeType };
}

function errorResponse(status: number, code: string, message: string): Response {
  return Response.json(
    { error: { message, type: "invalid_request_error", param: null, code } },
    { status },
  );
}
