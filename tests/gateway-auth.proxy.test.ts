import { describe, expect, test } from "bun:test";
import { promises as fs } from "node:fs";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { loadWrapperOptions } from "../src/proxy/config";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import { ResponseStore } from "../src/proxy/response-store";
import { createProxyApp } from "../src/proxy/server";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import { parseListenUrl } from "../src/proxy/utils";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";

describe("gateway authentication and task scope", () => {
  test("loads gateway token and task scope from CONFIG__ overrides", async () => {
    const previousToken = process.env.CONFIG__gatewayToken;
    const previousScope = process.env.CONFIG__taskScope;
    process.env.CONFIG__gatewayToken = "123456";
    process.env.CONFIG__taskScope = "general";
    try {
      const options = await loadWrapperOptions(process.cwd());
      expect(options.gatewayToken).toBe("123456");
      expect(options.taskScope).toBe("general");
    } finally {
      if (previousToken === undefined) {
        delete process.env.CONFIG__gatewayToken;
      } else {
        process.env.CONFIG__gatewayToken = previousToken;
      }
      if (previousScope === undefined) {
        delete process.env.CONFIG__taskScope;
      } else {
        process.env.CONFIG__taskScope = previousScope;
      }
    }
  });

  test("keeps requests working when no gateway token is configured", async () => {
    const calls = { create: 0, chat: 0 };
    const app = createApp(createOptions(), calls);

    const response = await post(app);

    expect(response.status).toBe(200);
    expect((await response.json()).status).toBe("completed");
    expect(calls.create).toBe(1);
    expect(calls.chat).toBe(1);
  });

  test("requires the configured runtime token on model-provider routes", async () => {
    const options = createOptions();
    options.gatewayToken = "seeded-runtime-token";
    const calls = { create: 0, chat: 0 };
    const app = createApp(options, calls);

    const missing = await post(app);
    const wrong = await post(app, "wrong-token");
    const correct = await post(app, "seeded-runtime-token");

    expect(missing.status).toBe(403);
    expect(wrong.status).toBe(403);
    expect((await missing.json()).error.code).toBe("permission_denied");
    expect((await wrong.json()).error.code).toBe("permission_denied");
    expect(correct.status).toBe(200);
    expect((await correct.json()).status).toBe("completed");
    expect(calls.create).toBe(1);
    expect(calls.chat).toBe(1);
  });

  test("requires the configured runtime token on trace retrieval", async () => {
    const options = createOptions();
    options.gatewayToken = "seeded-runtime-token";
    const app = createApp(options, { create: 0, chat: 0 });

    const missing = await app.fetch(
      new Request("http://localhost/__viz/traces/trace-1"),
    );
    const correct = await app.fetch(
      new Request("http://localhost/__viz/traces/trace-1", {
        headers: { "x-runtime-token": "seeded-runtime-token" },
      }),
    );

    expect(missing.status).toBe(403);
    expect(correct.status).toBe(404);
  });

  test("rejects general task scope on chat completions before upstream invocation", async () => {
    const options = createOptions();
    options.taskScope = "general";
    const calls = { create: 0, chat: 0 };
    const app = createApp(options, calls);

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          model: "m365-copilot",
          messages: [{ role: "user", content: "hello" }],
        }),
      }),
    );
    const body = (await response.json()) as JsonObject;

    expect(response.status).toBe(400);
    expect(body.error.code).toBe("capability_unavailable");
    expect(calls.create).toBe(0);
    expect(calls.chat).toBe(0);
  });

  test("leaves health and readiness unauthenticated", async () => {
    for (const gatewayToken of [null, "seeded-runtime-token"]) {
      const options = createOptions();
      options.gatewayToken = gatewayToken;
      const app = createApp(options, { create: 0, chat: 0 });

      const health = await app.fetch(new Request("http://localhost/healthz"));
      const ready = await app.fetch(new Request("http://localhost/readyz"));

      expect(health.status).toBe(200);
      expect(ready.status).toBe(200);
    }
  });

  test("only accepts loopback listen addresses", () => {
    expect(parseListenUrl("http://127.0.0.1:4000")).toEqual({
      hostname: "127.0.0.1",
      port: 4000,
    });
    expect(() => parseListenUrl("http://0.0.0.0:4000")).toThrow(
      "loopback",
    );
  });

  test("does not expose the runtime token in diagnostics or errors", async () => {
    const token = "seeded-runtime-token";
    const options = createOptions();
    options.gatewayToken = token;
    const app = createApp(options, { create: 0, chat: 0 });

    const health = await app.fetch(new Request("http://localhost/healthz"));
    const ready = await app.fetch(new Request("http://localhost/readyz"));
    const error = await post(app);

    expect(await health.text()).not.toContain(token);
    expect(await ready.text()).not.toContain(token);
    expect(await error.text()).not.toContain(token);

    const logPath = "tests/.gateway-auth-log-check";
    await fs.rm(logPath, { recursive: true, force: true });
    try {
      const logger = new DebugMarkdownLogger(
        { ...options, debugPath: logPath, logStdout: false },
        true,
      );
      await logger.logIncomingRequest(
        new Request("http://localhost/v1/responses", {
          headers: { "x-runtime-token": token },
        }),
        null,
      );
      const entries = await fs.readdir(logPath);
      const contents = await Promise.all(
        entries.map((entry) => fs.readFile(`${logPath}/${entry}`, "utf8")),
      );
      expect(contents.join("\n")).not.toContain(token);
    } finally {
      await fs.rm(logPath, { recursive: true, force: true });
    }
  });

  test("rejects general task scope before any upstream invocation", async () => {
    const options = createOptions();
    options.taskScope = "general";
    const calls = { create: 0, chat: 0 };
    const app = createApp(options, calls);

    const response = await post(app, undefined, { stream: true });
    const events = await collectEvents(response);
    const terminalEvents = events.filter((event) =>
      ["response.completed", "response.failed", "response.incomplete"].includes(
        String(event.type),
      ),
    );

    expect(response.status).toBe(200);
    expect(terminalEvents).toHaveLength(1);
    expect(terminalEvents[0]?.type).toBe("response.failed");
    const terminalResponse = terminalEvents[0]?.response as JsonObject;
    expect((terminalResponse.error as JsonObject).code).toBe(
      "capability_unavailable",
    );
    expect(calls.create).toBe(0);
    expect(calls.chat).toBe(0);
  });
});

function createApp(
  options: WrapperOptions,
  calls: { create: number; chat: number },
): ReturnType<typeof createProxyApp> {
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);
  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: `conv-${++calls.create}`,
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
      _payload: JsonObject,
    ): Promise<ChatResult> => {
      calls.chat += 1;
      return {
        isSuccess: true,
        statusCode: 200,
        responseJson: null,
        rawBody: "{}",
        assistantText: "hello",
        conversationId,
      };
    },
    chatOverStream: async (): Promise<Response> =>
      new Response(null, { status: 500 }),
  } as unknown as CopilotGraphClient;
  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv-substrate",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("not used");
    },
    chatStream: async (): Promise<ChatResult> => {
      throw new Error("not used");
    },
  } as unknown as CopilotSubstrateClient;
  const debugLogger = {
    logIncomingRequest: async () => {},
    logOutgoingResponse: async () => {},
    logUpstreamRequest: async () => {},
    logUpstreamResponse: async () => {},
    logSubstrateFrame: async () => {},
  } as unknown as DebugMarkdownLogger;
  const tokenProvider = {
    resolveAuthorizationHeader: async () => "******",
  } as unknown as ProxyTokenProvider;

  return createProxyApp({
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    tokenProvider,
  });
}

async function post(
  app: ReturnType<typeof createProxyApp>,
  gatewayToken?: string,
  overrides: JsonObject = {},
): Promise<Response> {
  return app.fetch(
    new Request("http://localhost/v1/responses", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-m365-transport": TransportNames.Graph,
        ...(gatewayToken ? { "x-runtime-token": gatewayToken } : {}),
      },
      body: JSON.stringify({
        model: "m365-copilot",
        input: "hello",
        ...overrides,
      }),
    }),
  );
}

async function collectEvents(response: Response): Promise<JsonObject[]> {
  const body = await response.text();
  const events: JsonObject[] = [];
  for (const block of body.split("\n\n")) {
    const dataLine = block
      .split("\n")
      .find((line) => line.startsWith("data: "));
    if (!dataLine) continue;
    const data = dataLine.slice("data: ".length);
    if (data === "[DONE]") continue;
    try {
      const parsed = JSON.parse(data) as unknown;
      if (
        parsed &&
        typeof parsed === "object" &&
        !Array.isArray(parsed)
      ) {
        events.push(parsed as JsonObject);
      }
    } catch {
      // Ignore malformed non-JSON SSE data in this test helper.
    }
  }
  return events;
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Error,
    openAiTransformMode: OpenAiTransformModes.Mapped,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Graph,
    graphBaseUrl: "https://graph.microsoft.com",
    createConversationPath: "/conversations",
    chatPathTemplate: "/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate: "/conversations/{conversationId}/stream",
    substrate: {
      hubPath: "wss://example.invalid",
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
      taskTimeoutSeconds: 900,
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
