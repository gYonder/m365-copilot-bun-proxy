import type { JsonObject, ParsedImageInput, SubstrateOptions } from "./types";

type ImageUploadResult =
  | { ok: true; annotations: JsonObject[] }
  | { ok: false; aborted: boolean };

export async function uploadSubstrateImages(
  images: readonly ParsedImageInput[],
  input: {
    rawToken: string;
    objectId: string;
    tenantId: string;
    conversationId: string;
    substrate: SubstrateOptions;
    signal?: AbortSignal | null;
    fetchImpl?: typeof fetch;
  },
): Promise<ImageUploadResult> {
  if (images.length === 0) {
    return { ok: true, annotations: [] };
  }

  const fetchImpl = input.fetchImpl ?? fetch;
  const annotations: JsonObject[] = [];
  for (let index = 0; index < images.length; index++) {
    if (input.signal?.aborted) {
      return { ok: false, aborted: true };
    }
    const result = await uploadOneImage(images[index]!, index, input, fetchImpl);
    if (!result.ok) {
      return result;
    }
    annotations.push(result.annotation);
  }
  return { ok: true, annotations };
}

async function uploadOneImage(
  image: ParsedImageInput,
  index: number,
  input: {
    rawToken: string;
    objectId: string;
    tenantId: string;
    conversationId: string;
    substrate: SubstrateOptions;
    signal?: AbortSignal | null;
  },
  fetchImpl: typeof fetch,
): Promise<
  | { ok: true; annotation: JsonObject }
  | { ok: false; aborted: boolean }
> {
  const body = new FormData();
  body.set("scenario", "UploadImage");
  body.set("conversationId", input.conversationId);
  body.set("FileBase64", image.dataUrl);
  body.set("optionsSets", "gptvnorm2048");

  const headers = new Headers({
    Authorization: `Bearer ${input.rawToken}`,
    Origin: input.substrate.origin || "https://m365.cloud.microsoft",
    "x-scenario": "OfficeWebIncludedCopilot",
    "x-variants": "feature.EnableImageSupportInUploadFile",
  });
  if (input.objectId && input.tenantId) {
    headers.set(
      "x-anchormailbox",
      `Oid:${input.objectId}@${input.tenantId}`,
    );
  }

  const timeoutMs = input.substrate.imageUploadTimeoutMs ?? 30_000;
  const controller = new AbortController();
  let timedOut = false;
  const timeout = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, timeoutMs);
  const onAbort = () => controller.abort();
  input.signal?.addEventListener("abort", onAbort, { once: true });
  try {
    const response = await fetchImpl(
      input.substrate.imageUploadUrl ??
        "https://substrate.office.com/m365Copilot/UploadFile",
      {
        method: "POST",
        headers,
        body,
        signal: controller.signal,
      },
    );
    if (response.status !== 200) {
      return { ok: false, aborted: false };
    }
    let responseJson: JsonObject | null = null;
    try {
      const parsed = (await response.json()) as unknown;
      responseJson =
        parsed && typeof parsed === "object" && !Array.isArray(parsed)
          ? (parsed as JsonObject)
          : null;
    } catch {
      return { ok: false, aborted: false };
    }
    if (
      responseJson?.result === null ||
      typeof responseJson?.result !== "object" ||
      Array.isArray(responseJson.result) ||
      (responseJson.result as JsonObject).value !== "Success"
    ) {
      return { ok: false, aborted: false };
    }
    const docId = stringValue(responseJson.docId);
    if (!docId) {
      return { ok: false, aborted: false };
    }
    const fileType = normalizeFileType(stringValue(responseJson.fileType));
    const fileName =
      stringValue(responseJson.fileName) || `image-${index + 1}.${fileType}`;
    return {
      ok: true,
      annotation: {
        id: docId,
        messageAnnotationType: "ImageFile",
        messageAnnotationMetadata: {
          "@type": "File",
          annotationType: "File",
          fileType,
          fileName,
        },
      },
    };
  } catch {
    // A timeout is an upload failure, not a caller cancellation: only an
    // externally signalled abort may be reported as cancelled.
    return { ok: false, aborted: !timedOut && input.signal?.aborted === true };
  } finally {
    clearTimeout(timeout);
    input.signal?.removeEventListener("abort", onAbort);
  }
}

function stringValue(value: unknown): string | null {
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

function normalizeFileType(value: string | null): string {
  const normalized = value?.replace(/^\.+/, "").trim();
  return normalized || "png";
}
