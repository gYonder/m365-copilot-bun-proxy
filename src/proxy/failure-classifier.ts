export const BridgeErrorCodes = {
  PermissionDenied: "permission_denied",
  AuthenticationRequired: "authentication_required",
  AuthorizationExpired: "authorization_expired",
  QuotaExhausted: "quota_exhausted",
  RateLimited: "rate_limited",
  CapabilityUnavailable: "capability_unavailable",
  SchemaInvalid: "schema_invalid",
  DuplicateSuppressed: "duplicate_suppressed",
  FallbackExhausted: "fallback_exhausted",
  ArtifactRejected: "artifact_rejected",
  PolicyBlocked: "policy_blocked",
  ImageUploadFailed: "image_upload_failed",
  TransportFailed: "transport_failed",
  ProviderDrift: "provider_drift",
  UpstreamTimeout: "upstream_timeout",
  AmbiguousCompletion: "ambiguous_completion",
  InternalError: "internal_error",
} as const;

export type BridgeErrorCode =
  (typeof BridgeErrorCodes)[keyof typeof BridgeErrorCodes];

export const INTERNAL_REASONS = [
  "invalid_request",
  "invalid_transport",
  "invalid_simulated_payload",
  "response_stream_error",
  "substrate_error",
  "graph_error",
  "missing_trace_id",
  "unknown_trace_id",
  "image_generation_disabled",
  "substrate_terminal_error",
  "substrate_incomplete_terminal",
  "image_generation_failed",
  "image_upload_failed",
  "gateway_caller_credential_rejected",
  "task_scope_not_supported",
  "interactive_sign_in_required",
  "authorization_expired",
  "conversation_quota_exhausted",
  "upstream_rate_limit",
  "capability_unavailable",
  "duplicate_tool_result_or_replay",
  "tool_round_repetition_bound_exhausted",
  "image_artifact_content_rejected",
  "upstream_policy_entitlement_or_feature_refusal",
  "provider_transport_or_artifact_retrieval_failed",
  "provider_drift",
  "upstream_timeout",
  "confab_recovery_exhausted",
  "partial_or_unprovable_completion",
  "unclassified_bridge_defect",
  "unknown_reason",
] as const;

export type InternalReason = (typeof INTERNAL_REASONS)[number];

export type BridgeCancellationReason = "client_aborted";

export type BridgeTerminal = "failed" | "incomplete";

export type BridgeFailure = {
  code: BridgeErrorCode;
  reason: InternalReason;
  message: string;
  retryable: boolean;
  terminal: BridgeTerminal;
};

export const INTERNAL_REASON_TO_ERROR_CODE: Record<
  InternalReason,
  BridgeErrorCode
> = {
  invalid_request: "schema_invalid",
  invalid_transport: "schema_invalid",
  invalid_simulated_payload: "schema_invalid",
  response_stream_error: "transport_failed",
  substrate_error: "transport_failed",
  graph_error: "transport_failed",
  missing_trace_id: "schema_invalid",
  unknown_trace_id: "schema_invalid",
  image_generation_disabled: "capability_unavailable",
  substrate_terminal_error: "transport_failed",
  substrate_incomplete_terminal: "ambiguous_completion",
  image_generation_failed: "artifact_rejected",
  image_upload_failed: "image_upload_failed",
  gateway_caller_credential_rejected: "permission_denied",
  task_scope_not_supported: "capability_unavailable",
  interactive_sign_in_required: "authentication_required",
  authorization_expired: "authorization_expired",
  conversation_quota_exhausted: "quota_exhausted",
  upstream_rate_limit: "rate_limited",
  capability_unavailable: "capability_unavailable",
  duplicate_tool_result_or_replay: "duplicate_suppressed",
  tool_round_repetition_bound_exhausted: "fallback_exhausted",
  image_artifact_content_rejected: "artifact_rejected",
  upstream_policy_entitlement_or_feature_refusal: "policy_blocked",
  provider_transport_or_artifact_retrieval_failed: "transport_failed",
  provider_drift: "provider_drift",
  upstream_timeout: "upstream_timeout",
  confab_recovery_exhausted: "ambiguous_completion",
  partial_or_unprovable_completion: "ambiguous_completion",
  unclassified_bridge_defect: "internal_error",
  unknown_reason: "internal_error",
};

const SAFE_MESSAGES: Record<BridgeErrorCode, string> = {
  permission_denied: "The request was not permitted.",
  authentication_required: "Interactive sign-in is required.",
  authorization_expired: "Authorization has expired.",
  quota_exhausted: "The provider quota has been exhausted.",
  rate_limited: "The provider is rate limiting requests.",
  capability_unavailable: "The requested capability is unavailable.",
  schema_invalid: "The request or tool schema is invalid.",
  duplicate_suppressed: "The duplicate request or tool result was suppressed.",
  fallback_exhausted: "No supported fallback remained.",
  artifact_rejected: "The provider rejected the artifact content.",
  image_upload_failed: "The provider rejected the image upload.",
  policy_blocked: "The provider blocked this request by policy.",
  transport_failed: "The provider transport failed.",
  provider_drift: "The provider protocol changed unexpectedly.",
  upstream_timeout: "The provider did not respond before the deadline.",
  ambiguous_completion: "The response ended before completion could be proven.",
  internal_error: "The bridge encountered an internal error.",
};

const RETRYABLE: Record<BridgeErrorCode, boolean> = {
  permission_denied: false,
  authentication_required: false,
  authorization_expired: true,
  quota_exhausted: false,
  rate_limited: true,
  capability_unavailable: false,
  schema_invalid: false,
  duplicate_suppressed: false,
  fallback_exhausted: false,
  artifact_rejected: false,
  image_upload_failed: false,
  policy_blocked: false,
  transport_failed: true,
  provider_drift: false,
  upstream_timeout: true,
  ambiguous_completion: false,
  internal_error: false,
};

export function classifyBridgeFailure(reason: InternalReason): BridgeFailure {
  const code = INTERNAL_REASON_TO_ERROR_CODE[reason];
  return {
    code,
    reason,
    message: SAFE_MESSAGES[code],
    retryable: RETRYABLE[code],
    terminal: code === "ambiguous_completion" ? "incomplete" : "failed",
  };
}

export function classifyUnknownReason(_raw: string): BridgeFailure {
  return classifyBridgeFailure("unknown_reason");
}
