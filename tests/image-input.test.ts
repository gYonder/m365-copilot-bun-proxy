import { describe, expect, test } from "bun:test";
import { parseImageInputs } from "../src/proxy/image-input";
import { tryParseOpenAiRequest } from "../src/proxy/request-parser";
import type { WrapperOptions } from "../src/proxy/types";

const png = "data:image/png;base64,AQID";
const jpeg = "data:image/jpeg;base64,BAUG";

describe("image input parsing", () => {
  test("accepts and decodes PNG and JPEG data URIs", () => {
    const result = parseImageInputs(
      [
        { type: "input_image", image_url: png },
        { type: "image_url", image_url: { url: jpeg } },
      ],
      {},
    );

    expect(result.ok).toBeTrue();
    if (result.ok) {
      expect(result.images.map((image) => image.byteLength)).toEqual([3, 3]);
      expect(result.images.map((image) => image.mimeType)).toEqual([
        "image/png",
        "image/jpeg",
      ]);
    }
  });

  test("supports a bare url field on image parts", () => {
    const result = parseImageInputs(
      [{ type: "image", url: png }],
      {},
    );

    expect(result.ok).toBeTrue();
    if (result.ok) {
      expect(result.images[0]?.dataUrl).toBe(png);
    }
  });

  test("rejects disallowed MIME types", () => {
    const result = parseImageInputs(
      [{ type: "image_url", image_url: "data:image/svg+xml;base64,AA==" }],
      {},
    );

    expect(result).toEqual({
      ok: false,
      reason: "unsupported_mime_type",
      message: "The image MIME type is not allowed.",
    });
  });

  test("rejects malformed base64 without exposing the payload", () => {
    const payload = "not-valid-base64";
    const result = parseImageInputs(
      [{ type: "input_image", image_url: `data:image/png;base64,${payload}` }],
      {},
    );

    expect(result.ok).toBeFalse();
    expect(JSON.stringify(result)).not.toContain(payload);
  });

  test("rejects remote URLs without making a network call", () => {
    let fetchCalls = 0;
    const result = parseImageInputs(
      [{ type: "image_url", image_url: "https://example.invalid/image.png" }],
      {},
    );

    expect(result.ok).toBeFalse();
    expect(result.reason).toBe("remote_url_rejected");
    expect(fetchCalls).toBe(0);
  });

  test("enforces per-image, total-byte, and image-count caps", () => {
    expect(
      parseImageInputs([{ type: "image", url: png }], {
        maxBytesPerImage: 2,
      }).reason,
    ).toBe("image_too_large");
    expect(
      parseImageInputs(
        [
          { type: "image", url: png },
          { type: "image", url: jpeg },
        ],
        { maxTotalBytes: 5 },
      ).reason,
    ).toBe("total_image_bytes_exceeded");
    expect(
      parseImageInputs(
        [
          { type: "image", url: png },
          { type: "image", url: png },
        ],
        { maxImages: 1 },
      ).reason,
    ).toBe("image_count_exceeded");
  });

  test("never emits the former attached-image text marker", () => {
    const parsed = tryParseOpenAiRequest(
      {
        model: "m365-copilot",
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: "Describe this." },
              { type: "input_image", image_url: png },
            ],
          },
        ],
      },
      createOptions(),
    );

    expect(parsed.ok).toBeTrue();
    if (parsed.ok) {
      expect(parsed.request.promptText).not.toContain("attached image");
      expect(parsed.request.promptText).not.toContain("image input(s)");
      expect(parsed.request.promptText).not.toContain(png);
      expect(parsed.request.images).toHaveLength(1);
    }
  });
  test("does not treat a non-image part with a url as an image", () => {
    const result = parseImageInputs(
      [{ type: "input_file", url: "https://example.invalid/spec.pdf" }],
      {},
    );

    expect(result).toEqual({ ok: true, images: [] });
  });

  test("keeps tool schemas that declare a url property intact", () => {
    const parsed = tryParseOpenAiRequest(
      {
        model: "m365-copilot",
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: "Describe this." },
              { type: "input_image", image_url: png },
            ],
          },
        ],
        tools: [
          {
            type: "function",
            function: {
              name: "fetch_page",
              parameters: {
                type: "object",
                properties: { url: { type: "string" } },
                required: ["url"],
              },
            },
          },
        ],
      },
      { ...createOptions(), openAiTransformMode: "simulated" } as WrapperOptions,
    );

    expect(parsed.ok).toBeTrue();
    if (parsed.ok) {
      expect(parsed.request.promptText).toContain("fetch_page");
      expect(parsed.request.promptText).toContain("url");
      expect(parsed.request.images).toHaveLength(1);
    }
  });
});

function createOptions(): WrapperOptions {
  return {
    openAiTransformMode: "mapped",
    defaultModel: "m365-copilot",
    defaultTimeZone: "UTC",
    maxAdditionalContextMessages: 16,
    substrate: {
      maxSendChars: 0,
      truncateBeforeSending: false,
      imageUploadEnabled: true,
      maxImagesPerRequest: 8,
      maxBytesPerImage: 10_485_760,
      maxTotalImageBytes: 20_971_520,
      allowedImageMimeTypes: ["image/png", "image/jpeg"],
    },
  } as WrapperOptions;
}
