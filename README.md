# m365-copilot-bun-proxy

TypeScript/Bun port of the original M365 Copilot Bun Proxy .NET proxy + CLI.

## Stack

- Bun runtime
- Hono for HTTP routing / reverse-proxy behavior
- Zod for configuration validation
- OpenTUI for interactive CLI chat UI

## Install

```bash
bun install
```

## Harness config examples

Example harness configuration files live in `harness-config-examples/`.

- `harness-config-examples/opencode.jsonc` provides a ready-to-use OpenCode config wired to this proxy on `http://localhost:4000/v1/`.
- `harness-config-examples/codex.toml` provides a Codex CLI profile using the Responses API wire format.
- Copy and adapt these examples for your local harness setup.

### Codex CLI

Start the proxy, then add the sample provider/profile from `harness-config-examples/codex.toml` to `~/.codex/config.toml` or pass the same values with `-c` overrides:

```bash
codex exec \
  --skip-git-repo-check \
  --cd /path/to/workspace \
  --sandbox workspace-write \
  -c 'approval_policy="never"' \
  -c 'model_providers.m365-copilot-bun-proxy={ name = "M365 Copilot Bun Proxy", base_url = "http://localhost:4000/v1", wire_api="responses" }' \
  -c 'model_provider="m365-copilot-bun-proxy"' \
  -m gpt-5.5 \
  'Read notes.md and write a short markdown summary to summary.md.'
```

The proxy returns both OpenAI-compatible `/v1/models` data and Codex-compatible model metadata. It also accepts both Chat Completions tool definitions (`function.name`) and Responses API tool definitions (`name` at the tool top level), which is required for Codex local shell/file tool calls. The native Codex slug `gpt-5.5` is accepted as an alias for the M365 Copilot GPT-5.5 reasoning tone.

## Run proxy

```bash
bun run start:proxy
```

To enable debug markdown logs (requires `debugPath` in config):

```bash
bun run start:proxy -- --debug
```

You can also pass an explicit value:

```bash
bun run start:proxy -- --debug=false
```

Logging level is configured with `logLevel` in `config.json` (default `info`):

- Each proxy startup with `--debug` writes logs into a timestamp-named session subfolder under `debugPath` (for example `logs/2026-02-25T16-58-11-123Z/`).
- Request logs (`incoming-request`, `request`, `substrate-request`) are always written when debug logging is enabled.
- Response log filtering by level:
- `trace`: includes `substrate-delta`
- `debug`: includes `substrate-response`, `response`, `response-headers`, and `outgoing-response`
- `info`: includes `outgoing-response`
- `warning`: includes `outgoing-response` only for HTTP 4xx status
- `error`: includes `outgoing-response` only for HTTP 5xx status

To capture outbound SSE payloads (for `text/event-stream` responses), set:

```json
{
  "logStreamingResponseBody": true
}
```

When enabled, and debug logging is active (`--debug`) with `logLevel` of `debug` or `trace`, the proxy writes `outgoing-stream-body` markdown logs containing the streamed SSE body.

Default listen URL is `http://localhost:4000`.

Configuration is loaded from `config.json` (and `config.{env}.json` when `NODE_ENV` is set).

Substrate settings are grouped under the `substrate` object in config (for example `substrate.hubPath`).

`ignoreIncomingAuthorizationHeader` controls whether inbound `Authorization` headers are used by the proxy. Default is `true`, which makes the proxy ignore incoming auth and use cached/auto-fetched tokens instead.

`playwrightBrowser` controls which Playwright browser is used when the proxy auto-acquires a token. Supported values: `edge` (default), `chrome`, `chromium`, `firefox`, `webkit` (`msedge` is also accepted as an alias for `edge`).

`temporaryChat` (default `true`) enables temporary-chat mode for Substrate by appending `disableMemory=1` to the websocket hub URL query string. This will prevent Copilot from showing your conversation history in the sidebar.

For proactive token refresh without opening a browser window, run:

```bash
bun src/cli/index.ts token fetch --headless --quiet
```

Headless token fetch uses the saved Playwright browser state, opens M365 Copilot in temporary-chat mode, sends a tiny message to trigger the Substrate websocket, and stores the captured token locally. If Microsoft requires interactive sign-in, rerun without `--headless` and complete the browser login.

`openAiTransformMode` controls how requests are translated for M365 Copilot:

- `simulated` (default): sends the full incoming OpenAI JSON payload with an
  endpoint-specific output contract. The proxy buffers the complete upstream
  turn, validates it before emitting output, and locally builds the OpenAI
  response or terminal failure.
- Tool-free simulated requests ask for direct assistant text and project that
  text into the requested OpenAI response shape. Strict whole-envelope
  validation remains mandatory whenever tools are available.
- `mapped`: uses the legacy request/response mapping logic.

`simulatedOutputProtocol` selects the tool-bearing Responses contract:

- `legacy` (default): requires the validated endpoint JSON envelope.
- `bridge_v1`: accepts exact bridge-owned frames for one final message, one
  function call, or one custom-tool call, with strict legacy JSON retained as a
  compatibility fallback. JSON response formats always use `legacy`.

Enable the V1 contract for a canary with:

```bash
CONFIG__simulatedOutputProtocol=bridge_v1 bun run start:proxy
```

The V1 frames start at byte zero:

```text
M365_FINAL_V1
<raw final text>

M365_FUNCTION_TOOL_CALL_V1
<exact tool name>
<one JSON argument object>

M365_CUSTOM_TOOL_CALL_V1
<exact tool name>
<raw input to EOF>
```

Recognized malformed frames are corrected or rejected; they are never accepted
as final prose. Function arguments, offered tool names, tool choice, schemas,
call IDs, and parallel policy still pass the same strict validator used by the
legacy protocol.

The legacy `substrate.earlyCompleteOnSimulatedPayload` and
`substrate.incrementalSimulatedContentStreaming` settings remain accepted for
configuration compatibility but are ignored. Simulated output is never emitted
incrementally or before whole-response validation; mapped streaming behavior is
unchanged.

Use `CONFIG__openAiTransformMode=mapped` if you need to revert to the legacy behavior.

You can override config values via env vars with the `CONFIG__` prefix, for example:

```bash
CONFIG__listenUrl=http://localhost:4010 bun run start:proxy
```

Example: force automatic token acquisition to use Chrome instead of Edge:

```bash
CONFIG__playwrightBrowser=chrome bun run start:proxy
```

To override nested values, use double underscores for each path segment, for example:

```bash
CONFIG__substrate__hubPath=wss://substrate.office.com/m365Copilot/Chathub bun run start:proxy
```

Sanitized bridge events can be persisted independently of debug request logs:

```json
{
  "observability": {
    "enabled": true,
    "logPath": "./logs/proxy-events.jsonl",
    "maxBytes": 5242880,
    "maxFiles": 3
  }
}
```

The JSONL log contains event names, counters, classifications, and sizes only.
Prompt and response bodies, authentication material, account identifiers, and
authenticated URLs are redacted. Rotation retains the active file plus numbered
archives up to `maxFiles`.

## API endpoints

- `POST /v1/chat/completions`
- `POST /openai/v1/chat/completions`
- `GET /v1/models`
- `GET /openai/v1/models`
- `POST /v1/responses`
- `POST /openai/v1/responses`
- `GET /v1/responses`
- `GET /openai/v1/responses`
- `GET /v1/responses/{response_id}`
- `GET /openai/v1/responses/{response_id}`
- `DELETE /v1/responses/{response_id}`
- `DELETE /openai/v1/responses/{response_id}`

## Available models

The proxy accepts any OpenAI-compatible `model` string, but for Substrate transport it maps known model IDs to a `tone` value in the outgoing websocket invocation payload.

`tone` is the Copilot UI option to pick a model type, out of:

- "Auto" => `magic`
- "Quick Response" => `Chat`
- "Think Deeper" => `Reasoning`
- "GPT5.2 Quick" => `Gpt_5_2_Chat`
- "GPT5.2 Think deeper" => `Gpt_5_2_Reasoning`
- "GPT5.4 Quick" => `Gpt_5_4_Chat`
- "GPT5.4 Think deeper" => `Gpt_5_4_Reasoning`
- "GPT5.5 Quick" => `Gpt_5_5_Chat`
- "GPT5.5 Think deeper" => `Gpt_5_5_Reasoning`

Model to Substrate `tone` mapping:

- `m365-copilot` -> `magic`
- `m365-copilot-auto` -> `magic`
- `m365-copilot-magic` -> `magic`
- Any unknown model value -> `magic`
- `m365-copilot-quick` -> `Chat`
- `m365-copilot-reasoning` -> `Reasoning`
- `m365-copilot-gpt5.2-quick` -> `Gpt_5_2_Chat`
- `m365-copilot-gpt5.2-reasoning` -> `Gpt_5_2_Reasoning`
- `m365-copilot-gpt5.4-quick` -> `Gpt_5_4_Chat`
- `m365-copilot-gpt5.4-reasoning` -> `Gpt_5_4_Reasoning`
- `m365-copilot-gpt5.5-quick` -> `Gpt_5_5_Chat`
- `m365-copilot-gpt5.5-reasoning` -> `Gpt_5_5_Reasoning`
- `gpt-5.5` -> `Gpt_5_5_Reasoning`

Notes:

- If `model` is omitted, the proxy uses `defaultModel` from config (defaults to `gpt-5.6-sol`).
- `GET /v1/models` (and `GET /openai/v1/models`) returns the full supported model list above.
- The GPT-5.5 reasoning tone (`Gpt_5_5_Reasoning`) and updated Substrate defaults are based on current Microsoft 365 Copilot web traffic captured in May 2026.

## Chat Completions Tool Calling

The proxy supports OpenAI-style `tools` and `tool_choice` for `POST /v1/chat/completions`.

Example request:

```bash
curl -s http://localhost:4000/v1/chat/completions \
  -H "Content-Type: application/json" \
  -H "x-m365-transport: substrate" \
  -d '{
    "model": "m365-copilot",
    "messages": [
      { "role": "user", "content": "What is the weather in London?" }
    ],
    "tools": [
      {
        "type": "function",
        "function": {
          "name": "get_weather",
          "description": "Lookup weather by city",
          "parameters": {
            "type": "object",
            "properties": {
              "city": { "type": "string" }
            },
            "required": ["city"]
          }
        }
      }
    ],
    "tool_choice": "auto"
  }'
```

Example tool-call response shape:

```json
{
  "id": "chatcmpl_...",
  "object": "chat.completion",
  "created": 1739986369,
  "model": "m365-copilot",
  "choices": [
    {
      "index": 0,
      "message": {
        "role": "assistant",
        "content": null,
        "tool_calls": [
          {
            "id": "call_...",
            "type": "function",
            "function": {
              "name": "get_weather",
              "arguments": "{\"city\":\"London\"}"
            }
          }
        ]
      },
      "finish_reason": "tool_calls"
    }
  ]
}
```

Strictness behavior:

- Tool-free simulated turns accept direct assistant text. Tool-bearing turns use
  the configured strict protocol described above.
- Strict JSON decoding first applies deterministic lexical repair only inside
  JSON strings: literal control characters are escaped, and invalid escape
  prefixes are preserved as literal backslashes. Truncation, quote insertion,
  brace balancing, surrounding prose, multiple values, and structural damage
  are never guessed or repaired.
- A malformed or invalid candidate receives one bounded protocol-correction
  turn. If that primary cycle is exhausted, the next sequential identical
  request receives one fresh regeneration cycle with an independent prompt and
  one correction. Later identical retries replay the cached failed terminal.
  This bounds a live failure episode to four upstream generations without
  durably storing failed response bodies.
- Tool choice, offered name and namespace, function/custom kind, argument schema,
  call IDs, and parallel-call policy are validated before any output is emitted.
  Mixed valid/invalid call batches are rejected atomically.
- Malformed-output telemetry contains only parse categories, lengths, booleans,
  root-shape metadata, and parser-provided offsets. It never records prompts,
  candidates, tool input, URLs, identifiers, or authentication material.

Input normalization notes:

- Validated function argument and custom-tool input bytes are preserved for
  downstream delivery; malformed input is rejected rather than repaired or
  evaluated.
- Assistant message content containing serialized `tool_calls` structures is preserved as tool-call context for downstream Copilot prompt construction.

## Responses API usage

Create response:

```bash
curl -s http://localhost:4000/v1/responses \
  -H "Content-Type: application/json" \
  -H "x-m365-transport: substrate" \
  -d '{
    "model": "m365-copilot",
    "input": "Write a TypeScript function that validates UUIDs."
  }'
```

Continue a conversation using `previous_response_id`:

```bash
curl -s http://localhost:4000/v1/responses \
  -H "Content-Type: application/json" \
  -H "x-m365-transport: substrate" \
  -d '{
    "model": "m365-copilot",
    "previous_response_id": "resp_abc123",
    "input": "Now add tests."
  }'
```

Streaming (`stream: true`) emits SSE events:

- `response.created`
- `response.in_progress`
- `response.output_item.added`
- `response.output_text.delta`
- `response.output_text.done`
- `response.output_item.done`
- `response.completed`
- `error` (SSE error event on stream failure)

By default, the proxy ignores inbound `Authorization` and attempts to use a cached token or auto-acquire one via Playwright for chat/responses requests.

The browser used for that auto-acquisition is controlled by `playwrightBrowser` in config (or `CONFIG__playwrightBrowser` in env).

To allow pass-through `Authorization` headers from clients, set:

```bash
CONFIG__ignoreIncomingAuthorizationHeader=false bun run start:proxy
```

## Build executable

```bash
bun run build
```

This produces a single-file executable in `dist/` and copies `config.json` alongside it.

## Run CLI

```bash
bun run cli -- help
bun run cli -- status
bun run cli -- chat
bun run cli -- chat --api responses
bun run cli -- token set --token "<jwt>"
bun src/cli/index.ts token fetch --headless --quiet
```

By default, CLI chat requests do not send an Authorization header. The proxy handles token acquisition when needed. Use `--token` or `YARPILOT_TOKEN` only when you want to force a specific token from the CLI.

In chat mode, the CLI supports these slash commands:

- `/status` (token + connection status)
- `/api` (show current API mode)
- `/api completions` or `/api responses` (toggle endpoint)
- `/token` (paste a new token)
- `/cleartoken` (clear cached token)
- `/exit` (quit)
