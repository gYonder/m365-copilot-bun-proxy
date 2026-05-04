import { promises as fs } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { chromium, firefox, webkit } from "playwright";

const SUBSTRATE_WS_PATTERN = /substrate\.office\.com\/m365Copilot\/Chathub/i;
const SUBSTRATE_WS_HOST_PATTERN = /(^|\.)substrate\.office\.com$/i;
const SUBSTRATE_WS_PATH_PATTERN = /\/m365Copilot\/Chathub\/?$/i;
const CHAT_URL = "https://m365.cloud.microsoft/chat/?auth=2";
const CHAT_URL_GLOB = "**/chat/**";
const LOGIN_HOST_PATTERN = /login\.(microsoftonline|live|microsoft)\.com/i;
const CHROMIUM_LAUNCH_ARGS = [
  "--no-first-run",
  "--no-default-browser-check",
  "--disable-search-engine-choice-screen",
];
const SUPPORTED_BROWSERS = new Set([
  "edge",
  "chrome",
  "chromium",
  "firefox",
  "webkit",
]);

const DEFAULT_TOKEN_TIMEOUT_MS = 120_000;
const DEFAULT_LOGIN_TIMEOUT_MS = 300_000;

if (isMainModule(process.argv[1])) {
  await runCli();
}

function isMainModule(entryPath) {
  if (!entryPath?.trim()) {
    return false;
  }
  try {
    return import.meta.url === pathToFileURL(entryPath).href;
  } catch {
    return false;
  }
}

async function runCli() {
  const parsed = parseArgs(process.argv.slice(2));
  const tokenPath = parsed["token-path"];
  const storageStatePath = parsed["storage-state-path"];
  const requestedBrowser = normalizeBrowserName(parsed.browser ?? "edge");
  const headless = parsed.headless !== undefined;
  const loginTimeoutMs = parsePositiveIntegerOption(
    parsed["login-timeout-ms"],
    DEFAULT_LOGIN_TIMEOUT_MS,
  );
  const tokenTimeoutMs = parsePositiveIntegerOption(
    parsed["token-timeout-ms"],
    DEFAULT_TOKEN_TIMEOUT_MS,
  );

  if (!tokenPath || !storageStatePath || !requestedBrowser) {
    const browserHelp = [...SUPPORTED_BROWSERS].join(", ");
    console.error(
      `Missing or invalid args. Required: --token-path <path> --storage-state-path <path> [--browser <${browserHelp}>]`,
    );
    process.exit(2);
  }

  await fetchTokenWithPlaywrightNode(tokenPath, storageStatePath, requestedBrowser, {
    headless,
    loginTimeoutMs,
    tokenTimeoutMs,
  });
}

async function fetchTokenWithPlaywrightNode(
  tokenPath,
  storageStatePath,
  browserName,
  options,
) {
  const mode = options.headless ? "headless" : "headed";
  console.log(
    `[playwright] Launching ${browserName} under Node.js (${mode})...`,
  );
  const browser = await launchBrowser(browserName, options.headless);
  const storageStateExists = await fileExists(storageStatePath);
  const context = await browser.newContext(
    storageStateExists ? { storageState: storageStatePath } : {},
  );
  await installSubstrateTemporaryChatShim(context);
  console.log(
    `[playwright] Browser launched (${storageStateExists ? "using saved storage state" : "fresh context"}).`,
  );

  try {
    const page = await context.newPage();
    await page.bringToFront().catch(() => {});
    console.log(`[playwright] Page URL: ${page.url()}`);
    const tokenCapture = captureSubstrateToken(page, options.tokenTimeoutMs);
    tokenCapture.promise.catch(() => {});

    try {
      console.log(`[playwright] Navigating to ${CHAT_URL}`);
      await page.goto(CHAT_URL, {
        waitUntil: "domcontentloaded",
        timeout: 60_000,
      });
      console.log(`[playwright] Landed on: ${page.url()}`);

      if (LOGIN_HOST_PATTERN.test(page.url())) {
        if (options.headless) {
          throw new Error(
            "Microsoft login is required; headless token capture cannot complete interactive sign-in.",
          );
        }
        console.log("[playwright] Login required - sign in in the browser window.");
        await page.waitForURL(CHAT_URL_GLOB, { timeout: options.loginTimeoutMs });
        console.log(`[playwright] Login complete: ${page.url()}`);
      } else {
        console.log("[playwright] Already logged in.");
      }

      await page
        .locator('[data-testid="newChatButton"], button[aria-label="New chat"]')
        .first()
        .click({ timeout: 5_000 })
        .catch(() => {});

      try {
        const editor = page.locator("#m365-chat-editor-target-element");
        await editor.waitFor({ state: "visible", timeout: 20_000 });
        console.log("[playwright] Sending message to trigger WebSocket...");
        await fillChatEditor(page, "Hi");
        await page.waitForTimeout(1_000);

        const sendButton = page
          .locator('button[title="Send"], [role="button"][title="Send"], [title="Send"]')
          .first();
        try {
          await sendButton.waitFor({ state: "visible", timeout: 10_000 });
          await sendButton.click({ timeout: 10_000 });
          console.log("[playwright] Send button clicked.");
        } catch {
          await page.keyboard.press("Enter");
          console.log("[playwright] Send button unavailable, submitted with Enter.");
        }
      } catch {
        console.log(
          "[playwright] Chat editor not found - waiting passively for WebSocket...",
        );
      }

      console.log(
        `[playwright] Waiting up to ${options.tokenTimeoutMs / 1000}s for token...`,
      );
      const rawToken = await tokenCapture.promise;
      console.log("[playwright] Token captured!");

      const expiresAtUtc = tryGetJwtExpiry(rawToken) ?? new Date(Date.now() + 3_600_000);
      await saveToken(tokenPath, rawToken, expiresAtUtc);
      await fs.mkdir(path.dirname(storageStatePath), { recursive: true });
      await context.storageState({ path: storageStatePath });
      console.log(`[playwright] Browser state saved: ${storageStatePath}`);
      console.log(`Token saved. Expires: ${expiresAtUtc.toISOString()}`);
    } finally {
      tokenCapture.cancel();
    }
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
  }
}

async function fillChatEditor(page, text) {
  const focused = await page.evaluate(() => {
    const editor = document.querySelector("#m365-chat-editor-target-element");
    if (!editor) {
      return false;
    }
    editor.focus();
    const target = editor.querySelector("p") ?? editor;
    const range = document.createRange();
    range.selectNodeContents(target);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    return true;
  });

  if (!focused) {
    throw new Error("M365 chat editor was not found.");
  }

  await page.keyboard.type(text, { delay: 10 });
}

async function installSubstrateTemporaryChatShim(context) {
  // Ensure token-fetch prompt does not get persisted in Copilot history.
  await context.addInitScript(({ hostPattern, pathPattern }) => {
    const substrateHostPattern = new RegExp(hostPattern, "i");
    const substrateHubPathPattern = new RegExp(pathPattern, "i");
    const OriginalWebSocket = window.WebSocket;

    const normalizeSubstrateHubUrl = (inputUrl) => {
      const raw = typeof inputUrl === "string" ? inputUrl : String(inputUrl);
      let parsed;
      try {
        parsed = new URL(raw, window.location.href);
      } catch {
        return raw;
      }

      const isSubstrateHub =
        substrateHostPattern.test(parsed.hostname) &&
        substrateHubPathPattern.test(parsed.pathname);
      if (!isSubstrateHub || parsed.searchParams.has("disableMemory")) {
        return raw;
      }

      parsed.searchParams.set("disableMemory", "1");
      return parsed.toString();
    };

    function WrappedWebSocket(url, protocols) {
      const nextUrl = normalizeSubstrateHubUrl(url);
      if (typeof protocols === "undefined") {
        return new OriginalWebSocket(nextUrl);
      }
      return new OriginalWebSocket(nextUrl, protocols);
    }

    WrappedWebSocket.prototype = OriginalWebSocket.prototype;
    Object.setPrototypeOf(WrappedWebSocket, OriginalWebSocket);
    for (const key of ["CONNECTING", "OPEN", "CLOSING", "CLOSED"]) {
      Object.defineProperty(WrappedWebSocket, key, {
        configurable: true,
        enumerable: true,
        value: OriginalWebSocket[key],
      });
    }

    Object.defineProperty(window, "WebSocket", {
      configurable: true,
      enumerable: false,
      writable: true,
      value: WrappedWebSocket,
    });
  }, {
    hostPattern: SUBSTRATE_WS_HOST_PATTERN.source,
    pathPattern: SUBSTRATE_WS_PATH_PATTERN.source,
  });
}

export function withDisableMemoryForSubstrateHubUrl(inputUrl, baseUrl) {
  const raw = typeof inputUrl === "string" ? inputUrl : String(inputUrl);
  let parsed;
  try {
    parsed = new URL(raw, baseUrl);
  } catch {
    return raw;
  }

  const isSubstrateHub =
    SUBSTRATE_WS_HOST_PATTERN.test(parsed.hostname) &&
    SUBSTRATE_WS_PATH_PATTERN.test(parsed.pathname);
  if (!isSubstrateHub || parsed.searchParams.has("disableMemory")) {
    return raw;
  }

  parsed.searchParams.set("disableMemory", "1");
  return parsed.toString();
}

async function launchBrowser(browserName, headless) {
  switch (browserName) {
    case "edge":
      try {
        return await chromium.launch({
          headless,
          channel: "msedge",
          args: CHROMIUM_LAUNCH_ARGS,
        });
      } catch {
        console.log(
          "[playwright] Edge channel unavailable, falling back to Chromium.",
        );
        return chromium.launch({
          headless,
          args: CHROMIUM_LAUNCH_ARGS,
        });
      }
    case "chrome":
      try {
        return await chromium.launch({
          headless,
          channel: "chrome",
          args: CHROMIUM_LAUNCH_ARGS,
        });
      } catch {
        console.log(
          "[playwright] Chrome channel unavailable, falling back to Chromium.",
        );
        return chromium.launch({
          headless,
          args: CHROMIUM_LAUNCH_ARGS,
        });
      }
    case "chromium":
      return chromium.launch({
        headless,
        args: CHROMIUM_LAUNCH_ARGS,
      });
    case "firefox":
      return firefox.launch({
        headless,
      });
    case "webkit":
      return webkit.launch({
        headless,
      });
    default:
      throw new Error(`Unsupported browser: ${String(browserName)}`);
  }
}

function captureSubstrateToken(page, timeoutMs = DEFAULT_TOKEN_TIMEOUT_MS) {
  let timer;
  let handler;
  let settled = false;

  const cleanup = () => {
    if (timer) clearTimeout(timer);
    if (handler) page.off("websocket", handler);
  };

  const promise = new Promise((resolve, reject) => {
    timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(
        new Error(
          `Timed out waiting for Substrate WebSocket after ${timeoutMs / 1000}s. Try running 'token fetch' again.`,
        ),
      );
    }, timeoutMs);

    handler = (ws) => {
      const url = ws.url();
      if (!SUBSTRATE_WS_PATTERN.test(url)) return;

      console.log("[playwright] Substrate WebSocket detected.");
      try {
        const token = new URL(url).searchParams.get("access_token");
        if (token) {
          if (settled) return;
          settled = true;
          cleanup();
          console.log("[playwright] access_token extracted.");
          resolve(token);
        }
      } catch {
        // Ignore parse failures from malformed websocket URLs.
      }
    };

    page.on("websocket", handler);
  });

  return {
    promise,
    cancel() {
      if (settled) return;
      settled = true;
      cleanup();
    },
  };
}

async function saveToken(filePath, token, expiresAtUtc) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(
    filePath,
    JSON.stringify(
      {
        token,
        expiresAtUtc: expiresAtUtc.toISOString(),
      },
      null,
      2,
    ),
    "utf8",
  );
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function tryGetJwtExpiry(token) {
  if (!token.trim()) {
    return null;
  }
  const parts = token.split(".");
  if (parts.length < 2) {
    return null;
  }

  try {
    const payload = Buffer.from(base64UrlNormalize(parts[1]), "base64").toString(
      "utf8",
    );
    const parsed = JSON.parse(payload);
    const expRaw = parsed.exp;
    const exp =
      typeof expRaw === "number"
        ? expRaw
        : typeof expRaw === "string"
          ? Number.parseInt(expRaw, 10)
          : Number.NaN;
    if (!Number.isFinite(exp)) {
      return null;
    }
    return new Date(exp * 1000);
  } catch {
    return null;
  }
}

function base64UrlNormalize(encoded) {
  const normalized = encoded.replaceAll("-", "+").replaceAll("_", "/");
  const padding = normalized.length % 4;
  return padding > 0
    ? normalized.padEnd(normalized.length + (4 - padding), "=")
    : normalized;
}

function parseArgs(args) {
  const options = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg.startsWith("--")) continue;
    const key = arg.slice(2);
    if (i + 1 < args.length && !args[i + 1].startsWith("--")) {
      options[key] = args[++i];
    } else {
      options[key] = "";
    }
  }
  return options;
}

function normalizeBrowserName(value) {
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return null;
  }
  const canonical = normalized === "msedge" ? "edge" : normalized;
  return SUPPORTED_BROWSERS.has(canonical) ? canonical : null;
}

function parsePositiveIntegerOption(value, defaultValue) {
  if (typeof value !== "string" || !value.trim()) {
    return defaultValue;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : defaultValue;
}
