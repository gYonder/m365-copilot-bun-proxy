import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";

const tempDirs: string[] = [];

afterEach(() => {
  while (tempDirs.length > 0) {
    const dir = tempDirs.pop();
    if (!dir) {
      continue;
    }
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("DebugMarkdownLogger substrate response logging", () => {
  test("debug mode emits structured frame JSON and omits metadata-only frames", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(createOptions(debugPath), true);

    const metadataOnlyFrame = {
      type: 1,
      target: "update",
      arguments: [{ requestId: "req-1", nonce: "n1" }],
    };
    const firstTextFrame = {
      type: 1,
      target: "update",
      arguments: [
        {
          requestId: "req-1",
          messages: [
            {
              author: "bot",
              messageId: "msg-1",
              text: "SECRET-CONVERSATION-CONTENT",
            },
          ],
        },
      ],
    };
    const completeJsonFrame = {
      type: 1,
      target: "update",
      arguments: [
        {
          requestId: "req-1",
          messages: [
            {
              author: "bot",
              messageId: "msg-1",
              text:
                "```json\n" +
                JSON.stringify(
                  {
                    id: "chatcmpl-sim-1",
                    object: "chat.completion",
                    choices: [
                      {
                        index: 0,
                        finish_reason: "stop",
                        message: { role: "assistant", content: "ok" },
                      },
                    ],
                  },
                  null,
                  2,
                ) +
                "\n```",
            },
          ],
        },
      ],
    };
    const terminalFrame = { type: 3, invocationId: "0" };

    const payload = [
      JSON.stringify(metadataOnlyFrame),
      JSON.stringify(firstTextFrame),
      JSON.stringify(completeJsonFrame),
      JSON.stringify(terminalFrame),
    ].join("\u001e");

    await logger.logSubstrateFrame(
      "wss://substrate.office.com/m365Copilot/Chathub",
      "response",
      `${payload}\u001e`,
    );

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-substrate-response.md"),
    );
    expect(files.length).toBe(1);

    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    const startMarker = "```json\n";
    const startIndex = content.indexOf(startMarker);
    const endIndex = content.lastIndexOf("\n```");
    expect(startIndex).toBeGreaterThanOrEqual(0);
    expect(endIndex).toBeGreaterThan(startIndex);
    const jsonText = content
      .slice(startIndex + startMarker.length, endIndex)
      .trim();
    const parsed = JSON.parse(jsonText) as JsonObject;

    expect(parsed.format).toBe("signalr-json-v1");
    expect(parsed.frameCount).toBe(4);
    expect(parsed.includedFrameCount).toBe(3);
    expect(parsed.omittedFrameCount).toBe(1);
    expect(Array.isArray(parsed.frames)).toBeTrue();

    const frames = parsed.frames as JsonObject[];
    expect(content.includes("SECRET-CONVERSATION-CONTENT")).toBeFalse();
    expect(content.includes("SECRET-CURSOR-CONTENT")).toBeFalse();
    const reasons = frames
      .map((frame) => String(frame.reason ?? ""))
      .filter((value) => value.length > 0);
    expect(reasons.includes("first_text")).toBeTrue();
    expect(reasons.includes("complete_markdown_json")).toBeTrue();
    expect(reasons.includes("terminal")).toBeTrue();
    for (const frame of frames) {
      expect("messagePreview" in frame).toBeFalse();
      expect("writeAtCursorPreview" in frame).toBeFalse();
    }
    const secretTextFrameLog = frames.find(
      (frame) => frame.messageTextLength === "SECRET-CONVERSATION-CONTENT".length,
    );
    expect(secretTextFrameLog?.messageTextLength).toBe(
      "SECRET-CONVERSATION-CONTENT".length,
    );
  });

  test("trace mode retains SignalR structure without conversation previews", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(
      createOptions(debugPath, LogLevels.Trace),
      true,
    );

    const payload = `${JSON.stringify({
      type: 1,
      target: "update",
      arguments: [
        {
          requestId: "req-trace",
          writeAtCursor: "TRACE-CURSOR-CONTENT",
          messages: [
            {
              author: "bot",
              messageId: "msg-trace",
              text: "TRACE-CONVERSATION-CONTENT",
            },
          ],
        },
      ],
    })}\u001e`;

    await logger.logSubstrateFrame(
      "wss://substrate.office.com/m365Copilot/Chathub",
      "request",
      payload,
    );

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-substrate-request.md"),
    );
    expect(files.length).toBe(1);
    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    expect(content.includes("TRACE-CONVERSATION-CONTENT")).toBeFalse();
    expect(content.includes("TRACE-CURSOR-CONTENT")).toBeFalse();
    expect(content.includes("\"messageTextLength\": 26")).toBeTrue();
    expect(content.includes("\"writeAtCursorLength\": 20")).toBeTrue();
    expect(content.includes("\"requestId\": \"req-trace\"")).toBeTrue();
    expect(content.includes("\"type\": 1")).toBeTrue();
  });

  test("trace mode writes simulated streaming diagnostics", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(
      createOptions(debugPath, LogLevels.Trace),
      true,
    );

    await logger.logSimulatedStreamingDiagnostics({
      completionId: "chatcmpl-test",
      outcome: "completed",
      parseAttemptCount: 3,
      parseSuccessCount: 1,
    });

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-simulated-streaming.md"),
    );
    expect(files.length).toBe(1);
    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    expect(content.includes("Simulated Streaming Diagnostics")).toBeTrue();
    expect(content.includes("\"parseAttemptCount\": 3")).toBeTrue();
  });

  test("debug mode does not write simulated streaming diagnostics", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(
      createOptions(debugPath, LogLevels.Debug),
      true,
    );

    await logger.logSimulatedStreamingDiagnostics({
      completionId: "chatcmpl-test",
      outcome: "completed",
    });

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-simulated-streaming.md"),
    );
    expect(files.length).toBe(0);
  });

  test("debug mode writes outgoing stream body when enabled", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const options = createOptions(debugPath, LogLevels.Debug);
    options.logStreamingResponseBody = true;
    const logger = new DebugMarkdownLogger(options, true);

    await logger.logOutgoingStreamBody(
      200,
      [["content-type", "text/event-stream"]],
      'event: response.completed\ndata: {"type":"response.completed"}\n\ndata: [DONE]\n\n',
    );

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-outgoing-stream-body.md"),
    );
    expect(files.length).toBe(1);
    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    expect(content.includes("Outgoing Stream Body")).toBeTrue();
    expect(content.includes("response.completed")).toBeTrue();
    expect(content.includes("[DONE]")).toBeTrue();
  });

  test("info mode does not write outgoing stream body", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const options = createOptions(debugPath, LogLevels.Info);
    options.logStreamingResponseBody = true;
    const logger = new DebugMarkdownLogger(options, true);

    await logger.logOutgoingStreamBody(
      200,
      [["content-type", "text/event-stream"]],
      'event: response.completed\ndata: {"type":"response.completed"}\n\ndata: [DONE]\n\n',
    );

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-outgoing-stream-body.md"),
    );
    expect(files.length).toBe(0);
  });

  test("normal logging omits request and response bodies", async () => {
    const debugPath = path.resolve("tests/.logger-body-check");
    rmSync(debugPath, { recursive: true, force: true });
    try {
      const logger = new DebugMarkdownLogger(
        createOptions(debugPath, LogLevels.Info),
        true,
      );
      const secret = "seeded-body-secret";
      await logger.logIncomingRequest(
        new Request("http://localhost/v1/responses", { method: "POST" }),
        JSON.stringify({ secret }),
      );
      await logger.logUpstreamRequest(
        "POST",
        "https://example.invalid/upstream",
        [],
        JSON.stringify({ secret }),
      );
      await logger.logOutgoingResponse(
        200,
        [["content-type", "application/json"]],
        JSON.stringify({ secret }),
      );

      const files = readdirSync(debugPath);
      const content = files
        .map((file) => readFileSync(path.join(debugPath, file), "utf8"))
        .join("\n");
      expect(content).not.toContain(secret);
    } finally {
      rmSync(debugPath, { recursive: true, force: true });
    }
  });
});

function createOptions(
  debugPath: string,
  logLevel: (typeof LogLevels)[keyof typeof LogLevels] = LogLevels.Debug,
): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath,
    logLevel,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Substrate,
    graphBaseUrl: "https://graph.microsoft.com",
    createConversationPath: "/beta/copilot/conversations",
    chatPathTemplate: "/beta/copilot/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate:
      "/beta/copilot/conversations/{conversationId}/chatOverStream",
    substrate: {
      hubPath: "wss://substrate.office.com/m365Copilot/Chathub",
      source: "officeweb",
      quoteSourceInQuery: true,
      scenario: "OfficeWebIncludedCopilot",
      origin: "https://m365.cloud.microsoft",
      product: "Office",
      agentHost: "Bizchat.FullScreen",
      licenseType: "Starter",
      agent: "web",
      variants: null,
      clientPlatform: "web",
      productThreadType: "Office",
      invocationTimeoutSeconds: 120,
      keepAliveSeconds: 15,
      optionsSets: [],
      allowedMessageTypes: [],
      invocationTarget: "chat",
      invocationType: 4,
      locale: "en-US",
      experienceType: "Default",
      entityAnnotationTypes: [],
      earlyCompleteOnSimulatedPayload: false,
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}
