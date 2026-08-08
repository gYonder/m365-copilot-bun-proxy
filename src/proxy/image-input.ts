import type { JsonObject, JsonValue, ParsedImageInput } from "./types";
import { isJsonObject, tryGetString } from "./utils";

type ImageInputRejectionReason =
  | "image_upload_disabled"
  | "image_url_missing"
  | "remote_url_rejected"
  | "invalid_data_uri"
  | "unsupported_mime_type"
  | "malformed_base64"
  | "empty_payload"
  | "image_count_exceeded"
  | "image_too_large"
  | "total_image_bytes_exceeded";

export type ImageInputRejection = {
  ok: false;
  reason: ImageInputRejectionReason;
  message: string;
};

export type ImageInputParseResult =
  | { ok: true; images: ParsedImageInput[] }
  | ImageInputRejection;

const DEFAULT_ALLOWED_MIME_TYPES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
];

export function parseImageInputs(
  contentNode: JsonValue | undefined,
  limits: {
    enabled?: boolean;
    maxImages?: number;
    maxBytesPerImage?: number;
    maxTotalBytes?: number;
    allowedMimeTypes?: string[];
  } = {},
): ImageInputParseResult {
  const parts = collectImageParts(contentNode);
  if (parts.length === 0) {
    return { ok: true, images: [] };
  }
  if (limits.enabled === false) {
    return reject(
      "image_upload_disabled",
      "Image input is disabled by the current configuration.",
    );
  }

  const maxImages = limits.maxImages ?? 8;
  if (parts.length > maxImages) {
    return reject(
      "image_count_exceeded",
      `The request contains too many images; the maximum is ${maxImages}.`,
    );
  }

  const maxBytesPerImage = limits.maxBytesPerImage ?? 10_485_760;
  const maxTotalBytes = limits.maxTotalBytes ?? 20_971_520;
  const allowedMimeTypes = new Set(
    (limits.allowedMimeTypes ?? DEFAULT_ALLOWED_MIME_TYPES).map((mimeType) =>
      mimeType.split(";", 1)[0].trim().toLowerCase(),
    ),
  );
  const images: ParsedImageInput[] = [];
  let totalBytes = 0;

  for (const part of parts) {
    const imageUrl = extractImageUrl(part);
    if (!imageUrl) {
      return reject("image_url_missing", "An image input is missing its URL.");
    }
    const parsed = parseDataUrl(imageUrl, allowedMimeTypes);
    if (!parsed.ok) {
      return parsed;
    }
    if (parsed.image.byteLength > maxBytesPerImage) {
      return reject(
        "image_too_large",
        `An image exceeds the per-image byte limit of ${maxBytesPerImage}.`,
      );
    }
    totalBytes += parsed.image.byteLength;
    if (totalBytes > maxTotalBytes) {
      return reject(
        "total_image_bytes_exceeded",
        `The images exceed the total byte limit of ${maxTotalBytes}.`,
      );
    }
    images.push(parsed.image);
  }

  return { ok: true, images };
}

function collectImageParts(contentNode: JsonValue | undefined): JsonObject[] {
  if (isImagePart(contentNode)) {
    return [contentNode];
  }
  if (!Array.isArray(contentNode)) {
    return [];
  }
  return contentNode.filter(isImagePart);
}

function isImagePart(value: JsonValue | undefined): value is JsonObject {
  if (!isJsonObject(value)) {
    return false;
  }
  const type = (tryGetString(value, "type") ?? "").trim().toLowerCase();
  // A bare `url` field is deliberately not enough to claim a part is an image:
  // non-image parts also carry urls, and misreading one fails the whole request.
  return (
    type === "input_image" ||
    type === "image_url" ||
    type === "image" ||
    value.image_url !== undefined
  );
}

function extractImageUrl(part: JsonObject): string | null {
  const imageUrl = part.image_url;
  if (typeof imageUrl === "string" && imageUrl.trim()) {
    return imageUrl.trim();
  }
  if (isJsonObject(imageUrl)) {
    const nestedUrl = tryGetString(imageUrl, "url");
    if (nestedUrl?.trim()) {
      return nestedUrl.trim();
    }
  }
  const directUrl = tryGetString(part, "url");
  return directUrl?.trim() || null;
}

function parseDataUrl(
  dataUrl: string,
  allowedMimeTypes: Set<string>,
):
  | { ok: true; image: ParsedImageInput }
  | ImageInputRejection {
  if (!dataUrl.toLowerCase().startsWith("data:")) {
    return reject(
      "remote_url_rejected",
      "Only inline data URI image inputs are supported.",
    );
  }

  const commaIndex = dataUrl.indexOf(",");
  if (commaIndex <= 5) {
    return reject("invalid_data_uri", "The image data URI is malformed.");
  }
  const metadata = dataUrl.slice(5, commaIndex);
  const payload = dataUrl.slice(commaIndex + 1);
  const metadataParts = metadata.split(";");
  const mimeType = metadataParts[0]?.trim().toLowerCase() ?? "";
  const hasBase64Marker = metadataParts
    .slice(1)
    .some((part) => part.trim().toLowerCase() === "base64");
  if (!mimeType || !hasBase64Marker) {
    return reject(
      "invalid_data_uri",
      "Image data URIs must use base64 encoding.",
    );
  }
  if (!allowedMimeTypes.has(mimeType)) {
    return reject(
      "unsupported_mime_type",
      "The image MIME type is not allowed.",
    );
  }
  if (!payload) {
    return reject("empty_payload", "The image data URI payload is empty.");
  }

  const decoded = decodeBase64(payload);
  if (!decoded) {
    return reject("malformed_base64", "The image data URI payload is invalid.");
  }
  return {
    ok: true,
    image: {
      dataUrl,
      mimeType,
      byteLength: decoded.byteLength,
    },
  };
}

function decodeBase64(payload: string): Uint8Array | null {
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(payload)) {
    return null;
  }
  const firstPadding = payload.indexOf("=");
  if (firstPadding >= 0 && firstPadding < payload.length - 2) {
    return null;
  }
  if (payload.length % 4 === 1) {
    return null;
  }
  try {
    const decoded = Buffer.from(payload, "base64");
    const normalizedPayload = payload.replace(/=+$/, "");
    const normalizedEncoded = decoded.toString("base64").replace(/=+$/, "");
    return normalizedPayload === normalizedEncoded ? decoded : null;
  } catch {
    return null;
  }
}

function reject(
  reason: ImageInputRejectionReason,
  message: string,
): ImageInputRejection {
  return { ok: false, reason, message };
}
