import { describe, expect, test } from "bun:test";
import {
  BridgeErrorCodes,
  classifyBridgeFailure,
  classifyUnknownReason,
  INTERNAL_REASON_TO_ERROR_CODE,
  INTERNAL_REASONS,
  type BridgeErrorCode,
  type InternalReason,
} from "../src/proxy/failure-classifier";

const REAL_REASON_MAPPINGS = [
  ["invalid_request", "schema_invalid"],
  ["invalid_transport", "schema_invalid"],
  ["invalid_simulated_payload", "schema_invalid"],
  ["response_stream_error", "transport_failed"],
  ["substrate_error", "transport_failed"],
  ["graph_error", "transport_failed"],
  ["missing_trace_id", "schema_invalid"],
  ["unknown_trace_id", "schema_invalid"],
  ["image_generation_disabled", "capability_unavailable"],
  ["substrate_terminal_error", "transport_failed"],
  ["substrate_incomplete_terminal", "ambiguous_completion"],
  ["confab_recovery_exhausted", "ambiguous_completion"],
  ["image_generation_failed", "artifact_rejected"],
  ["image_upload_failed", "image_upload_failed"],
] as const satisfies ReadonlyArray<readonly [InternalReason, BridgeErrorCode]>;

const SHARED_REASON_MAPPINGS = {
  permission_denied: "gateway_caller_credential_rejected",
  authentication_required: "interactive_sign_in_required",
  authorization_expired: "authorization_expired",
  quota_exhausted: "conversation_quota_exhausted",
  rate_limited: "upstream_rate_limit",
  capability_unavailable: "capability_unavailable",
  schema_invalid: "invalid_request",
  duplicate_suppressed: "duplicate_tool_result_or_replay",
  fallback_exhausted: "tool_round_repetition_bound_exhausted",
  artifact_rejected: "image_generation_failed",
  image_upload_failed: "image_upload_failed",
  policy_blocked: "upstream_policy_entitlement_or_feature_refusal",
  transport_failed: "substrate_error",
  provider_drift: "provider_drift",
  upstream_timeout: "upstream_timeout",
  ambiguous_completion: "partial_or_unprovable_completion",
  internal_error: "unclassified_bridge_defect",
} as const satisfies Record<BridgeErrorCode, InternalReason>;

describe("bridge failure classifier", () => {
  test("maps every real internal reason code explicitly", () => {
    for (const [reason, expectedCode] of REAL_REASON_MAPPINGS) {
      expect(classifyBridgeFailure(reason).code).toBe(expectedCode);
    }

    expect(classifyBridgeFailure("substrate_terminal_error").terminal).toBe(
      "failed",
    );
    expect(
      classifyBridgeFailure("substrate_incomplete_terminal").terminal,
    ).toBe("incomplete");
  });

  test("keeps the internal reason union exhaustively mapped", () => {
    const mapping: Record<InternalReason, BridgeErrorCode> =
      INTERNAL_REASON_TO_ERROR_CODE;

    expect(Object.keys(mapping).sort()).toEqual([...INTERNAL_REASONS].sort());
    for (const reason of INTERNAL_REASONS) {
      expect(mapping[reason]).toBeTypeOf("string");
      expect(classifyBridgeFailure(reason).reason).toBe(reason);
    }
  });

  test("maps every shared error code to one terminal kind", () => {
    const failures = Object.entries(SHARED_REASON_MAPPINGS).map(
      ([code, reason]) => {
        const failure = classifyBridgeFailure(reason);
        expect(failure.code).toBe(code);
        return failure;
      },
    );

    expect(failures).toHaveLength(17);
    expect(
      failures.filter((failure) => failure.terminal === "incomplete"),
    ).toHaveLength(1);
    expect(
      failures.find((failure) => failure.code === "ambiguous_completion")
        ?.terminal,
    ).toBe("incomplete");
  });

  test("does not infer a code from unknown upstream prose", () => {
    const failure = classifyUnknownReason(
      "timeout quota prompt=never-share token=seeded-secret",
    );

    expect(failure).toEqual({
      code: "internal_error",
      reason: "unknown_reason",
      message: "The bridge encountered an internal error.",
      retryable: false,
      terminal: "failed",
    });
    expect(failure.message).not.toContain("seeded-secret");
    expect(failure.message).not.toContain("never-share");
  });

  test("keeps retryability explicit for non-retryable failures", () => {
    const nonRetryable = [
      "schema_invalid",
      "quota_exhausted",
      "capability_unavailable",
      "duplicate_suppressed",
      "fallback_exhausted",
      "permission_denied",
      "policy_blocked",
      "artifact_rejected",
    ];

    for (const code of nonRetryable) {
      const reason = SHARED_REASON_MAPPINGS[code as BridgeErrorCode];
      expect(classifyBridgeFailure(reason).retryable).toBeFalse();
    }
  });
});
