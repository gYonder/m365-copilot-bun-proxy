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

// Microsoft identity/SSO cookie domains. These are small and are what lets a
// saved session silently re-authenticate, so they are always preserved.
const IDENTITY_COOKIE_DOMAIN_PATTERN =
  /(^|\.)(login\.microsoftonline\.com|login\.microsoft\.com|login\.live\.com|microsoftonline\.com|microsoftazuread-sso\.com)$/i;
// m365.cloud.microsoft returns a "Bad Request" error page (HTTP 400) when the
// request Cookie header is too large. Playwright's persisted storageState can
// accumulate the chunked OhpAuth*/OhpToken* application cookies across renewals
// until the jar (~40 KB) crosses that limit and blocks token renewal entirely.
// When the non-identity (application) cookies exceed this budget we drop them
// and rely on the identity cookies to mint a fresh, correctly sized session.
const OVERSIZED_APP_COOKIE_BYTES = 24_000;
const HEADER_TOO_LARGE_TITLE_PATTERN =
  /bad request|request too long|header too (large|long)|cookie too large/i;

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
  const initialStorageState = await loadSanitizedStorageState(storageStatePath);
  const context = await browser.newContext(
    initialStorageState ? { storageState: initialStorageState } : {},
  );
  await installSubstrateTemporaryChatShim(context);
  console.log(
    `[playwright] Browser launched (${initialStorageState ? "using saved storage state" : "fresh context"}).`,
  );

  try {
    const page = await context.newPage();
    await page.bringToFront().catch(() => {});
    console.log(`[playwright] Page URL: ${page.url()}`);
    const tokenCapture = captureSubstrateToken(page, options.tokenTimeoutMs);
    tokenCapture.promise.catch(() => {});

    try {
      console.log(`[playwright] Navigating to ${CHAT_URL}`);
      let response = await page.goto(CHAT_URL, {
        waitUntil: "domcontentloaded",
        timeout: 60_000,
      });
      console.log(`[playwright] Landed on: ${page.url()}`);

      if (await isHeaderTooLargeError(page, response)) {
        console.log(
          "[playwright] M365 returned a 'request header too large' error page; the saved session cookies are oversized. Dropping application cookies and re-authenticating with the sign-in cookies.",
        );
        await dropOversizedAppCookies(context);
        response = await page.goto(CHAT_URL, {
          waitUntil: "domcontentloaded",
          timeout: 60_000,
        });
        console.log(`[playwright] Reloaded after cookie reset: ${page.url()}`);
      }

      if (LOGIN_HOST_PATTERN.test(page.url())) {
        await completeSignIn(page, options);
      } else if (await isHeaderTooLargeError(page, response)) {
        throw new Error(
          "M365 still returns 'Bad Request - request header too large' after clearing application cookies. Run 'codex-m365-yolo --runtime-auth' and sign in when Edge opens to reset the saved session.",
        );
      } else {
        console.log("[playwright] Already logged in.");
      }

      await maybeHandleAuthDialog(page);
      await persistBrowserState(context, storageStatePath, "chat landing");

      const storageToken = await waitForStoredSubstrateToken(
        page,
        options.tokenTimeoutMs,
      );
      if (storageToken) {
        const jwtExpiry = tryGetJwtExpiry(storageToken.token);
        if (jwtExpiry) {
          console.log("[playwright] Token recovered from browser storage.");
          const expiresAtUtc = storageToken.expiresAtUtc ?? jwtExpiry;
          await saveToken(tokenPath, storageToken.token, expiresAtUtc, {
            oid: storageToken.oid ?? null,
            tid: storageToken.tid ?? null,
          });
          await persistBrowserState(context, storageStatePath, "storage token");
          console.log(`Token saved. Expires: ${expiresAtUtc.toISOString()}`);
          return;
        }

        console.log(
          "[playwright] Browser storage exposed a non-JWT substrate token; keeping browser state and waiting for a websocket token instead.",
        );
      }

      const storedClaims = storageToken
        ? {
            oid: storageToken.oid ?? null,
            tid: storageToken.tid ?? null,
          }
        : await extractStoredSubstrateAccountClaims(page);

      await waitForResourceAuthentication(
        context,
        page,
        options.tokenTimeoutMs,
      );

      console.log(
        "[playwright] No browser-stored substrate token found yet; falling back to WebSocket capture.",
      );

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
        const submitted = await trySubmitTriggerMessage(page);
        if (!submitted) {
          console.log(
            "[playwright] Could not verify trigger message submission; waiting passively for WebSocket...",
          );
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

      const jwtClaims = tryReadJwtClaims(rawToken);
      const oid =
        typeof jwtClaims?.oid === "string" && jwtClaims.oid.trim()
          ? jwtClaims.oid.trim()
          : storedClaims?.oid ?? null;
      const tid =
        typeof jwtClaims?.tid === "string" && jwtClaims.tid.trim()
          ? jwtClaims.tid.trim()
          : storedClaims?.tid ?? null;
      if (!jwtClaims && (!oid || !tid)) {
        throw new Error(
          "Captured websocket token is not a JWT and no reusable oid/tid metadata was available from browser state.",
        );
      }

      const expiresAtUtc = tryGetJwtExpiry(rawToken) ?? new Date(Date.now() + 3_600_000);
      await saveToken(tokenPath, rawToken, expiresAtUtc, {
        oid,
        tid,
      });
      await persistBrowserState(context, storageStatePath, "token capture");
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
  const editorLocator = page.locator("#m365-chat-editor-target-element");
  await editorLocator.click({ timeout: 10_000 });

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

  // Lexical ignores Playwright's key-by-key typing in the current M365
  // composer, while insertText dispatches the input event that activates Send.
  await page.keyboard.insertText(text);
  await page.waitForFunction(
    (expectedText) => {
      const editor = document.querySelector("#m365-chat-editor-target-element");
      return (editor?.textContent ?? "").includes(String(expectedText));
    },
    text,
    { timeout: 5_000 },
  );
}

async function maybeHandleAuthDialog(page, timeoutMs = 5_000) {
  const continueButton = page.getByRole("button", { name: /^continue$/i }).first();
  try {
    await continueButton.waitFor({ state: "visible", timeout: timeoutMs });
  } catch {
    return false;
  }

  await continueButton.click({ timeout: 10_000 });
  console.log("[playwright] Authentication dialog dismissed with Continue.");
  await continueButton.waitFor({ state: "hidden", timeout: 15_000 }).catch(() => {});
  await page.waitForTimeout(3_000);
  return true;
}

async function waitForResourceAuthentication(context, page, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let sawAuthentication = false;
  let reportedWaiting = false;
  let settledSince = null;

  while (Date.now() < deadline) {
    const continueButton = page
      .getByRole("button", { name: /^continue$/i })
      .first();
    const continueVisible = await continueButton
      .isVisible()
      .catch(() => false);
    if (continueVisible) {
      sawAuthentication = true;
      settledSince = null;
      await continueButton.click({ timeout: 10_000 }).catch(() => {});
    }

    const authenticationPages = context.pages().filter((candidatePage) => {
      if (candidatePage === page || candidatePage.isClosed()) {
        return false;
      }
      try {
        const hostname = new URL(candidatePage.url()).hostname.toLowerCase();
        return (
          hostname === "login.microsoftonline.com" ||
          hostname.endsWith(".login.microsoftonline.com")
        );
      } catch {
        return false;
      }
    });

    if (authenticationPages.length > 0) {
      sawAuthentication = true;
      settledSince = null;
      if (!reportedWaiting) {
        console.log(
          "[playwright] Microsoft resource authentication is open in Edge. Complete the approval there to continue.",
        );
        reportedWaiting = true;
      }
    } else if (!continueVisible) {
      if (!sawAuthentication) {
        return;
      }
      settledSince ??= Date.now();
      if (Date.now() - settledSince >= 3_000) {
        console.log("[playwright] Microsoft resource authentication completed.");
        return;
      }
    }

    await page.waitForTimeout(500);
  }

  throw new Error(
    `Timed out waiting for Microsoft resource authentication after ${timeoutMs / 1000}s.`,
  );
}

async function waitForStoredSubstrateToken(page, timeoutMs) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    await maybeHandleAuthDialog(page, 500);
    const token = await extractStoredSubstrateToken(page);
    if (token) {
      return token;
    }
    await page.waitForTimeout(500);
  }
  return null;
}

async function extractStoredSubstrateToken(page) {
  return page.evaluate(() => {
    const now = Date.now();
    const storages = [window.localStorage, window.sessionStorage];
    const preferredKeyPatterns = [
      /https:\/\/substrate\.office\.com\/\.default\|\|$/i,
      /https:\/\/substrate\.office\.com\/search\/\.default\|\|$/i,
      /https:\/\/substrate\.office\.com\/sydney\/\.default\|\|$/i,
      /https:\/\/substrate\.office\.com/i,
    ];

    const parseDateValue = (value) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        return value > 10_000_000_000 ? value : value * 1000;
      }
      if (typeof value !== "string" || !value.trim()) {
        return null;
      }
      const asNumber = Number(value);
      if (Number.isFinite(asNumber)) {
        return asNumber > 10_000_000_000 ? asNumber : asNumber * 1000;
      }
      const parsed = Date.parse(value);
      return Number.isFinite(parsed) ? parsed : null;
    };

    const parseJwtExpiry = (token) => {
      const parts = token.split(".");
      if (parts.length < 2) {
        return null;
      }
      try {
        const normalized = parts[1]
          .replace(/-/g, "+")
          .replace(/_/g, "/")
          .padEnd(parts[1].length + ((4 - (parts[1].length % 4)) % 4), "=");
        const payload = JSON.parse(atob(normalized));
        const exp = payload?.exp;
        if (typeof exp === "number" && Number.isFinite(exp)) {
          return exp * 1000;
        }
        if (typeof exp === "string" && Number.isFinite(Number(exp))) {
          return Number(exp) * 1000;
        }
      } catch {
        // Ignore invalid JWTs.
      }
      return null;
    };

    const claimSources = [];
    for (const storage of storages) {
      for (const key of Object.keys(storage)) {
        if (!/active-account-filters|account/i.test(key)) {
          continue;
        }
        try {
          const parsed = JSON.parse(storage.getItem(key) ?? "{}");
          const oid =
            typeof parsed.localAccountId === "string" && parsed.localAccountId.trim()
              ? parsed.localAccountId.trim()
              : typeof parsed.homeAccountId === "string" && parsed.homeAccountId.includes(".")
                ? parsed.homeAccountId.split(".", 1)[0].trim()
                : null;
          const tid =
            typeof parsed.tenantId === "string" && parsed.tenantId.trim()
              ? parsed.tenantId.trim()
              : typeof parsed.realm === "string" && parsed.realm.trim()
                ? parsed.realm.trim()
                : null;
          if (oid && tid) {
            claimSources.push({ oid, tid });
          }
        } catch {
          // Ignore malformed account metadata.
        }
      }
    }
    const preferredClaims =
      claimSources.find((candidate) => candidate.oid && candidate.tid) ?? null;

    const candidates = [];
    for (const storage of storages) {
      const keys = Object.keys(storage).filter(
        (key) => key.includes("accesstoken") && key.includes("substrate.office.com"),
      );
      for (const matchingKey of keys) {
        try {
          const parsed = JSON.parse(storage.getItem(matchingKey) ?? "{}");
          const token =
            typeof parsed.secret === "string"
              ? parsed.secret.trim()
              : typeof parsed.data === "string"
                ? parsed.data.trim()
                : "";
          if (token.length > 100) {
            const explicitExpiry =
              parseDateValue(parsed.expiresOn) ??
              parseDateValue(parsed.expires_on) ??
              parseDateValue(parsed.expiresAtUtc) ??
              parseJwtExpiry(token);
            const updatedAt =
              parseDateValue(parsed.lastUpdatedAt) ??
              parseDateValue(parsed.cachedAt) ??
              0;
            const preference =
              preferredKeyPatterns.findIndex((pattern) => pattern.test(matchingKey));
            candidates.push({
              token,
              expiresAtMs: explicitExpiry,
              updatedAtMs: updatedAt,
              preference: preference === -1 ? preferredKeyPatterns.length : preference,
              oid: preferredClaims?.oid ?? null,
              tid: preferredClaims?.tid ?? null,
            });
          }
        } catch {
          // Try the next candidate.
        }
      }
    }

    const unexpired = candidates
      .filter((candidate) => candidate.expiresAtMs && candidate.expiresAtMs > now + 60_000)
      .sort(
        (left, right) =>
          left.preference - right.preference ||
          right.expiresAtMs - left.expiresAtMs ||
          right.updatedAtMs - left.updatedAtMs,
      );
    if (unexpired.length > 0) {
      return {
        token: unexpired[0].token,
        expiresAtUtc: new Date(unexpired[0].expiresAtMs).toISOString(),
        oid: unexpired[0].oid,
        tid: unexpired[0].tid,
      };
    }

    const fallback = candidates.sort(
      (left, right) =>
        left.preference - right.preference ||
        right.updatedAtMs - left.updatedAtMs ||
        (right.expiresAtMs ?? 0) - (left.expiresAtMs ?? 0),
    )[0];
    if (!fallback) {
      return null;
    }

    return {
      token: fallback.token,
      expiresAtUtc: fallback.expiresAtMs
        ? new Date(fallback.expiresAtMs).toISOString()
        : null,
      oid: fallback.oid,
      tid: fallback.tid,
    };
  });
}

async function extractStoredSubstrateAccountClaims(page) {
  return page.evaluate(() => {
    const storages = [window.localStorage, window.sessionStorage];
    for (const storage of storages) {
      for (const key of Object.keys(storage)) {
        if (!/active-account-filters|account/i.test(key)) {
          continue;
        }
        try {
          const parsed = JSON.parse(storage.getItem(key) ?? "{}");
          const oid =
            typeof parsed.localAccountId === "string" && parsed.localAccountId.trim()
              ? parsed.localAccountId.trim()
              : typeof parsed.homeAccountId === "string" && parsed.homeAccountId.includes(".")
                ? parsed.homeAccountId.split(".", 1)[0].trim()
                : null;
          const tid =
            typeof parsed.tenantId === "string" && parsed.tenantId.trim()
              ? parsed.tenantId.trim()
              : typeof parsed.realm === "string" && parsed.realm.trim()
                ? parsed.realm.trim()
                : null;
          if (oid && tid) {
            return { oid, tid };
          }
        } catch {
          // Ignore malformed account metadata.
        }
      }
    }
    return null;
  });
}

async function trySubmitTriggerMessage(page) {
  await page.waitForTimeout(750);

  const sendButtonLocators = [
    page.getByRole("button", { name: /^send$/i }).first(),
    page
      .locator(
        [
          'button[aria-label="Send"]',
          '[role="button"][aria-label="Send"]',
          'button[title="Send"]',
          '[role="button"][title="Send"]',
          '[data-testid*="send"]',
        ].join(", "),
      )
      .first(),
  ];

  for (const locator of sendButtonLocators) {
    try {
      await locator.waitFor({ state: "visible", timeout: 5_000 });
      await locator.click({ timeout: 10_000 });
      if (await waitForEditorToClear(page, 2_500)) {
        console.log("[playwright] Trigger message submitted with Send button.");
        return true;
      }
    } catch {
      // Try the next submission method.
    }
  }

  const keyAttempts = [
    { keys: "Enter", label: "Enter" },
    { keys: "Meta+Enter", label: "Meta+Enter" },
    { keys: "Control+Enter", label: "Control+Enter" },
  ];

  for (const attempt of keyAttempts) {
    try {
      await page.keyboard.press(attempt.keys);
      if (await waitForEditorToClear(page, 2_500)) {
        console.log(
          `[playwright] Trigger message submitted with ${attempt.label}.`,
        );
        return true;
      }
    } catch {
      // Try the next submission method.
    }
  }

  return false;
}

async function waitForEditorToClear(page, timeoutMs) {
  try {
    await page.waitForFunction(
      () => {
        const editor = document.querySelector("#m365-chat-editor-target-element");
        return !!editor && !(editor.textContent ?? "").trim();
      },
      undefined,
      { timeout: timeoutMs },
    );
    return true;
  } catch {
    return false;
  }
}

async function completeSignIn(page, options) {
  if (options.headless) {
    // A saved SSO session can still complete a silent redirect back to the app
    // without any interaction. Give it a bounded window before giving up so a
    // headless refresh keeps working whenever the sign-in cookies are valid.
    console.log("[playwright] On Microsoft sign-in; waiting for silent SSO...");
    try {
      await page.waitForURL(CHAT_URL_GLOB, {
        timeout: Math.min(options.loginTimeoutMs, 25_000),
      });
      console.log(`[playwright] Silent SSO complete: ${page.url()}`);
      return;
    } catch {
      throw new Error(
        "Microsoft login is required; headless token capture cannot complete interactive sign-in.",
      );
    }
  }
  console.log("[playwright] Login required - sign in in the browser window.");
  await page.waitForURL(CHAT_URL_GLOB, { timeout: options.loginTimeoutMs });
  console.log(`[playwright] Login complete: ${page.url()}`);
}

async function loadSanitizedStorageState(storageStatePath) {
  if (!(await fileExists(storageStatePath))) {
    return null;
  }
  let state;
  try {
    state = JSON.parse(await fs.readFile(storageStatePath, "utf8"));
  } catch {
    console.log(
      "[playwright] Saved browser state was unreadable; starting from a fresh context.",
    );
    return null;
  }
  if (!state || typeof state !== "object" || !Array.isArray(state.cookies)) {
    return null;
  }

  const nowSeconds = Date.now() / 1000;
  const liveCookies = state.cookies.filter(
    (cookie) =>
      !(
        typeof cookie?.expires === "number" &&
        cookie.expires > 0 &&
        cookie.expires < nowSeconds
      ),
  );

  const appCookieBytes = sumCookieBytes(
    liveCookies.filter((cookie) => !isIdentityCookie(cookie)),
  );
  if (appCookieBytes > OVERSIZED_APP_COOKIE_BYTES) {
    const identityCookies = liveCookies.filter(isIdentityCookie);
    console.log(
      `[playwright] Saved application cookies are oversized (${appCookieBytes} bytes); keeping ${identityCookies.length} sign-in cookies so silent SSO can issue a fresh, correctly sized session.`,
    );
    return { cookies: identityCookies, origins: [] };
  }

  return { cookies: liveCookies, origins: state.origins ?? [] };
}

function isIdentityCookie(cookie) {
  const domain = String(cookie?.domain ?? "").replace(/^\./, "");
  return IDENTITY_COOKIE_DOMAIN_PATTERN.test(domain);
}

function sumCookieBytes(cookies) {
  return cookies.reduce(
    (total, cookie) =>
      total +
      String(cookie?.name ?? "").length +
      String(cookie?.value ?? "").length,
    0,
  );
}

async function dropOversizedAppCookies(context) {
  const cookies = await context.cookies();
  const identityCookies = cookies.filter(isIdentityCookie);
  await context.clearCookies();
  if (identityCookies.length > 0) {
    await context.addCookies(identityCookies);
  }
  console.log(
    `[playwright] Cleared ${cookies.length - identityCookies.length} application cookies; kept ${identityCookies.length} sign-in cookies.`,
  );
}

async function isHeaderTooLargeError(page, response) {
  // Only the m365.cloud.microsoft app origin produces the oversized-cookie
  // "Bad Request" page. On the login host we are mid sign-in, not in error.
  if (LOGIN_HOST_PATTERN.test(page.url())) {
    return false;
  }
  const status =
    response && typeof response.status === "function" ? response.status() : null;
  if (status === 400 || status === 431 || status === 494) {
    return true;
  }
  let title = "";
  try {
    title = (await page.title()) ?? "";
  } catch {
    title = "";
  }
  return HEADER_TOO_LARGE_TITLE_PATTERN.test(title);
}

async function persistBrowserState(context, storageStatePath, reason) {
  await fs.mkdir(path.dirname(storageStatePath), { recursive: true });
  await context.storageState({ path: storageStatePath });
  console.log(`[playwright] Browser state saved (${reason}): ${storageStatePath}`);
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
        } else {
          console.log(
            "[playwright] Substrate WebSocket detected, but no access_token was present in the URL.",
          );
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

async function saveToken(filePath, token, expiresAtUtc, metadata = {}) {
  const directory = path.dirname(filePath);
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const temporaryPath = path.join(
    directory,
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.tmp`,
  );
  const content = JSON.stringify(
    {
      token,
      expiresAtUtc: expiresAtUtc.toISOString(),
      ...(metadata.oid ? { oid: metadata.oid } : {}),
      ...(metadata.tid ? { tid: metadata.tid } : {}),
    },
    null,
    2,
  );
  try {
    await fs.writeFile(temporaryPath, content, { encoding: "utf8", mode: 0o600 });
    await fs.chmod(temporaryPath, 0o600);
    await fs.rename(temporaryPath, filePath);
    await fs.chmod(filePath, 0o600);
  } catch (error) {
    await fs.unlink(temporaryPath).catch(() => {});
    throw error;
  }
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
  const parsed = tryReadJwtClaims(token);
  if (!parsed) {
    return null;
  }
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
}

function tryReadJwtClaims(token) {
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
    return JSON.parse(payload);
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
