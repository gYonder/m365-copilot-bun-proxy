import { describe, expect, test } from "bun:test";
import {
  ImageGenerationError,
  ImageGenerationService,
  type ImageGenerationTransport,
} from "../src/proxy/image-generation";
import type { ImageGenerationOptions } from "../src/proxy/types";

const options = (
  overrides: Partial<ImageGenerationOptions> = {},
): ImageGenerationOptions => ({
  enabled: true,
  maxPromptChars: 100,
  maxImages: 2,
  maxArtifactBytes: 16,
  timeoutMs: 1_000,
  concurrencyLimit: 1,
  allowedMimeTypes: ["image/png"],
  ...overrides,
});

const request = (body: unknown, signal?: AbortSignal) =>
  new Request("http://localhost/v1/images/generations", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });

const transport = (
  generate: ImageGenerationTransport["generate"],
): ImageGenerationTransport => ({ generate });

async function errorCode(response: Response): Promise<string> {
  const body = (await response.json()) as { error: { code: string } };
  return body.error.code;
}

describe("isolated image generation endpoint service", () => {
  test("is disabled unless explicitly configured", async () => {
    const service = new ImageGenerationService(
      options({ enabled: false }),
      { resolveDesignerAuthorizationHeader: async () => "Bearer unused" },
      transport(async () => []),
    );
    const response = await service.create(request({ prompt: "unused" }));
    expect(response.status).toBe(404);
    expect(await errorCode(response)).toBe("image_generation_disabled");
  });

  test("returns base64 without exposing an authenticated artifact URL", async () => {
    const signedUrl =
      "https://artifact.example.invalid/image.png?signature=sensitive";
    const service = new ImageGenerationService(
      options(),
      { resolveDesignerAuthorizationHeader: async () => "Bearer designer-token" },
      transport(async (authorization, prompt, count) => {
        expect(authorization).toBe("Bearer designer-token");
        expect(prompt).toBe("draw a tree");
        expect(count).toBe(1);
        return [{ url: signedUrl, revisedPrompt: "a green tree" }];
      }),
      async (_url, init) => {
        expect(init?.redirect).toBe("manual");
        return new Response(new Uint8Array([1, 2, 3]), {
          headers: { "content-type": "image/png" },
        });
      },
    );
    const response = await service.create(
      request({ prompt: "draw a tree", response_format: "b64_json" }),
    );
    expect(response.status).toBe(200);
    const raw = await response.text();
    expect(raw).not.toContain(signedUrl);
    expect(raw).not.toContain("signature=sensitive");
    const body = JSON.parse(raw) as {
      data: Array<{ b64_json: string; revised_prompt: string }>;
    };
    expect(body.data).toEqual([
      { b64_json: "AQID", revised_prompt: "a green tree" },
    ]);
  });

  test("rejects prompt and count limits before acquiring a token", async () => {
    let tokenCalls = 0;
    const service = new ImageGenerationService(
      options({ maxPromptChars: 4, maxImages: 1 }),
      {
        resolveDesignerAuthorizationHeader: async () => {
          tokenCalls += 1;
          return "Bearer unused";
        },
      },
      transport(async () => []),
    );
    const longPrompt = await service.create(request({ prompt: "12345" }));
    const excessiveCount = await service.create(request({ prompt: "ok", n: 2 }));
    expect(await errorCode(longPrompt)).toBe("prompt_too_long");
    expect(await errorCode(excessiveCount)).toBe("invalid_image_count");
    expect(tokenCalls).toBe(0);
  });

  test("requires independent Designer authorization", async () => {
    const service = new ImageGenerationService(
      options(),
      { resolveDesignerAuthorizationHeader: async () => null },
      transport(async () => []),
    );
    const response = await service.create(request({ prompt: "tree" }));
    expect(response.status).toBe(401);
    expect(await errorCode(response)).toBe("designer_auth_unavailable");
  });

  test("preserves typed quota errors", async () => {
    const service = new ImageGenerationService(
      options(),
      { resolveDesignerAuthorizationHeader: async () => "Bearer designer" },
      transport(async () => {
        throw new ImageGenerationError(
          429,
          "image_quota_exceeded",
          "Image quota exceeded.",
        );
      }),
    );
    const response = await service.create(request({ prompt: "tree" }));
    expect(response.status).toBe(429);
    expect(await errorCode(response)).toBe("image_quota_exceeded");
  });

  for (const scenario of [
    {
      name: "redirect",
      response: new Response(null, {
        status: 302,
        headers: { location: "https://other.invalid" },
      }),
      code: "artifact_redirect_rejected",
    },
    {
      name: "unexpected MIME type",
      response: new Response("not an image", {
        headers: { "content-type": "text/html" },
      }),
      code: "unexpected_artifact_type",
    },
    {
      name: "oversized content",
      response: new Response(new Uint8Array(17), {
        headers: { "content-type": "image/png" },
      }),
      code: "artifact_too_large",
    },
    {
      name: "expired artifact URL",
      response: new Response("expired", { status: 403 }),
      code: "artifact_retrieval_failed",
    },
  ]) {
    test("rejects " + scenario.name, async () => {
      const service = new ImageGenerationService(
        options(),
        { resolveDesignerAuthorizationHeader: async () => "Bearer designer" },
        transport(async () => [
          { url: "https://artifact.example.invalid/image" },
        ]),
        async () => scenario.response.clone(),
      );
      const response = await service.create(request({ prompt: "tree" }));
      expect(response.status).toBe(502);
      expect(await errorCode(response)).toBe(scenario.code);
    });
  }

  test("cancels generation when the client aborts", async () => {
    const controller = new AbortController();
    const service = new ImageGenerationService(
      options(),
      { resolveDesignerAuthorizationHeader: async () => "Bearer designer" },
      transport(async (_authorization, _prompt, _count, signal) => {
        const aborted = new Promise<void>((resolve) =>
          signal.addEventListener("abort", () => resolve(), { once: true }),
        );
        controller.abort();
        await aborted;
        throw new DOMException("aborted", "AbortError");
      }),
    );
    const response = await service.create(
      request({ prompt: "tree" }, controller.signal),
    );
    expect(response.status).toBe(499);
    expect(await errorCode(response)).toBe("image_generation_cancelled");
  });
});
