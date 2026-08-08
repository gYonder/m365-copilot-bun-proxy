import { parseArgs } from "util";
import { promises as fs } from "node:fs";
import path from "node:path";
import { CopilotGraphClient, CopilotSubstrateClient } from "./clients";
import { loadWrapperOptions } from "./config";
import { ConversationStore } from "./conversation-store";
import { DebugMarkdownLogger } from "./logger";
import { DurableStateStore } from "./durable-state";
import { ResponseStore } from "./response-store";
import { createProxyApp } from "./server";
import { ProxyTokenProvider } from "./token-provider";
import { ProxyVizTraceStore } from "./viz-trace-store";
import type { WrapperOptions } from "./types";
import { parseListenUrl } from "./utils";
import { SubstrateSessionStore } from "./substrate-session-store";
import {
  DesignerSubstrateImageTransport,
  ImageGenerationService,
} from "./image-generation";
import { BridgeObservability } from "./observability";

const loadedOptions = await loadWrapperOptions(process.cwd());
const debugEnabled = parseDebugFlag();
const options = withSessionDebugPath(loadedOptions, debugEnabled);
if (debugEnabled && options.debugPath?.trim()) {
  await fs.mkdir(options.debugPath, { recursive: true });
}
const debugLogger = new DebugMarkdownLogger(options, debugEnabled);
const observability = new BridgeObservability();
const graphClient = new CopilotGraphClient(options, debugLogger);
const durableState = new DurableStateStore();
const substrateSessionStore = new SubstrateSessionStore(
  options.conversationTtlMinutes,
  undefined,
  durableState,
);
const substrateClient = new CopilotSubstrateClient(
  options,
  debugLogger,
  substrateSessionStore,
  undefined,
  undefined,
  undefined,
  observability,
);
const conversationStore = new ConversationStore(options, durableState, observability);
const responseStore = new ResponseStore(options, durableState, observability);
const vizTraceStore = new ProxyVizTraceStore(options.conversationTtlMinutes * 60);
const tokenProvider = new ProxyTokenProvider({
  ignoreIncomingAuthorizationHeader: options.ignoreIncomingAuthorizationHeader,
  playwrightBrowser: options.playwrightBrowser,
  msalAuthEnabled: options.msalAuth,
  observability,
});
const imageGenerationService = new ImageGenerationService(
  options.imageGeneration,
  tokenProvider,
  new DesignerSubstrateImageTransport(options, debugLogger),
  undefined,
  observability,
);

const app = createProxyApp({
  options,
  debugLogger,
  graphClient,
  substrateClient,
  conversationStore,
  responseStore,
  tokenProvider,
  vizTraceStore,
  imageGenerationService,
  observability,
});

const listen = parseListenUrl(options.listenUrl);
const debugLogPath =
  debugEnabled && options.debugPath?.trim()
    ? path.resolve(options.debugPath)
    : null;
const server = Bun.serve({
  idleTimeout: 0, // Disable idle timeout to allow long-running requests
  hostname: listen.hostname,
  port: listen.port,
  fetch: app.fetch,
});

process.stdout.write(
  `m365-copilot-bun-proxy listening on http://${server.hostname}:${server.port} (from ${options.listenUrl})\n`,
);
process.stdout.write(
  debugEnabled
    ? `Debugging: enabled (level: ${options.logLevel}, logs at ${debugLogPath})\n`
    : `Debugging: disabled (configured level: ${options.logLevel})\n`,
);

setInterval(() => undefined, 86_400_000);
await new Promise<never>(() => {});

function parseDebugFlag(): boolean {
  const { values } = parseArgs({
    args: Bun.argv,
    options: {
      debug: {
        type: "boolean",
        short: "d",
      },
    },
    strict: true,
    allowPositionals: true,
  });

  return values.debug ?? false;
}

function withSessionDebugPath(
  options: WrapperOptions,
  debugEnabled: boolean,
): WrapperOptions {
  if (!debugEnabled || !options.debugPath?.trim()) {
    return options;
  }

  const basePath = options.debugPath.trim();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

  return {
    ...options,
    debugPath: path.join(basePath, timestamp),
  };
}
