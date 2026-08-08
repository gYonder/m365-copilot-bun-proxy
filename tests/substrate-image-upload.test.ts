import { describe, expect, test } from "bun:test";
import { uploadSubstrateImages } from "../src/proxy/substrate-image-upload";
import type { ParsedImageInput, SubstrateOptions } from "../src/proxy/types";

const image: ParsedImageInput = {
  dataUrl: "data:image/png;base64,AQID",
  mimeType: "image/png",
  byteLength: 3,
};

describe("Substrate image upload", () => {
  test("builds the exact annotation shape and multipart fields", async () => {
    let capturedRequest: Request | null = null;
    const result = await uploadSubstrateImages([image], {
      rawToken: "token",
      objectId: "oid",
      tenantId: "tid",
      conversationId: "conversation-1",
      substrate: createSubstrate(),
      fetchImpl: async (input, init) => {
        capturedRequest = new Request(input, init);
        return successResponse({
          docId: "doc-1",
          fileName: "upload.png",
          fileType: ".png",
        });
      },
    });

    expect(result).toEqual({
      ok: true,
      annotations: [
        {
          id: "doc-1",
          messageAnnotationType: "ImageFile",
          messageAnnotationMetadata: {
            "@type": "File",
            annotationType: "File",
            fileType: "png",
            fileName: "upload.png",
          },
        },
      ],
    });
    expect(capturedRequest).not.toBeNull();
    const form = await capturedRequest!.formData();
    expect(form.get("scenario")).toBe("UploadImage");
    expect(form.get("optionsSets")).toBe("gptvnorm2048");
    expect(form.get("FileBase64")).toBe(image.dataUrl);
    expect(form.get("conversationId")).toBe("conversation-1");
    expect(capturedRequest!.headers.get("authorization")).toBe("Bearer token");
    expect(capturedRequest!.headers.get("origin")).toBe(
      "https://m365.cloud.microsoft",
    );
    expect(capturedRequest!.headers.get("x-scenario")).toBe(
      "OfficeWebIncludedCopilot",
    );
    expect(capturedRequest!.headers.get("x-variants")).toBe(
      "feature.EnableImageSupportInUploadFile",
    );
    expect(capturedRequest!.headers.get("x-anchormailbox")).toBe(
      "Oid:oid@tid",
    );
  });

  test("omits x-anchormailbox when oid or tid is absent", async () => {
    let capturedRequest: Request | null = null;
    await uploadSubstrateImages([image], {
      rawToken: "token",
      objectId: "",
      tenantId: "tid",
      conversationId: "conversation-1",
      substrate: createSubstrate(),
      fetchImpl: async (input, init) => {
        capturedRequest = new Request(input, init);
        return successResponse({ docId: "doc-1", fileType: "png" });
      },
    });

    expect(capturedRequest!.headers.has("x-anchormailbox")).toBeFalse();
  });

  test("requires HTTP 200 and result.value Success", async () => {
    const unsuccessfulBody = await uploadSubstrateImages([image], {
      ...uploadInput(),
      fetchImpl: async () =>
        successResponse({ result: { value: "Failure" }, docId: "doc-1" }),
    });
    const non200 = await uploadSubstrateImages([image], {
      ...uploadInput(),
      fetchImpl: async () =>
        new Response("not-success", { status: 201 }),
    });

    expect(unsuccessfulBody.ok).toBeFalse();
    expect(non200.ok).toBeFalse();
  });

  test("requires a non-empty docId", async () => {
    const result = await uploadSubstrateImages([image], {
      ...uploadInput(),
      fetchImpl: async () =>
        successResponse({ result: { value: "Success" }, docId: "" }),
    });

    expect(result.ok).toBeFalse();
  });

  test("treats transport errors as upload failures", async () => {
    const result = await uploadSubstrateImages([image], {
      ...uploadInput(),
      fetchImpl: async () => {
        throw new Error(image.dataUrl);
      },
    });

    expect(result).toEqual({ ok: false, aborted: false });
    expect(JSON.stringify(result)).not.toContain(image.dataUrl);
  });

  test("fails closed without returning partial annotations", async () => {
    let calls = 0;
    const result = await uploadSubstrateImages([image, image], {
      ...uploadInput(),
      fetchImpl: async () => {
        calls++;
        return calls === 1
          ? successResponse({ result: { value: "Success" }, docId: "doc-1" })
          : new Response("failure", { status: 500 });
      },
    });

    expect(calls).toBe(2);
    expect(result).toEqual({ ok: false, aborted: false });
  });

  test("reports a timeout as a failure rather than a cancellation", async () => {
    const result = await uploadSubstrateImages([image], {
      ...uploadInput(),
      substrate: {
        ...createSubstrate(),
        imageUploadTimeoutMs: 5,
      } as SubstrateOptions,
      fetchImpl: (_input, init) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () =>
            reject(new Error("aborted")),
          );
        }),
    });

    expect(result).toEqual({ ok: false, aborted: false });
  });

  test("reports an externally cancelled upload as aborted", async () => {
    const controller = new AbortController();
    const result = await uploadSubstrateImages([image], {
      ...uploadInput(),
      signal: controller.signal,
      fetchImpl: (_input, init) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () =>
            reject(new Error("aborted")),
          );
          controller.abort();
        }),
    });

    expect(result).toEqual({ ok: false, aborted: true });
  });
});

function uploadInput() {
  return {
    rawToken: "token",
    objectId: "oid",
    tenantId: "tid",
    conversationId: "conversation-1",
    substrate: createSubstrate(),
  };
}

function createSubstrate(): SubstrateOptions {
  return {
    origin: "https://m365.cloud.microsoft",
    imageUploadUrl: "https://substrate.office.com/m365Copilot/UploadFile",
    imageUploadTimeoutMs: 30_000,
  } as SubstrateOptions;
}

function successResponse(body: Record<string, unknown>): Response {
  return new Response(
    JSON.stringify({
      result: { value: "Success" },
      ...body,
    }),
    { status: 200, headers: { "content-type": "application/json" } },
  );
}
