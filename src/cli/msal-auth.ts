import { promises as fs } from "node:fs";
import path from "node:path";
import * as msal from "@azure/msal-node";
import type { BrowserContextOptions } from "playwright";
import { getBrowserStatePath, getTokenPath } from "./token-helpers";
import type { PlaywrightBrowserName } from "./playwright-token";

// First-party public-client app id that M365 Copilot ("Sydney") web uses. The
// authorization-code + PKCE flow against it yields substrate access tokens
// directly, so token renewal becomes a silent refresh-token call instead of a
// full browser cookie replay (which is what overflowed request headers).
const CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const AUTHORITY = "https://login.microsoftonline.com/common";
const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = [
  "https://substrate.office.com/sydney/M365Chat.Read",
  "https://substrate.office.com/sydney/sydney.readwrite",
];

export type MsalTokenResult = {
  token: string;
  expiresAtUtc: Date;
  oid: string | null;
  tid: string | null;
};

export type MsalAcquireOptions = {
  browserStatePath?: string;
  browser?: PlaywrightBrowserName;
  allowInteractive?: boolean;
  headless?: boolean;
  quiet?: boolean;
};

async function getMsalCachePath(): Promise<string> {
  const tokenPath = await getTokenPath();
  return path.join(path.dirname(tokenPath), "msal-cache.json");
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function decodeJwtClaims(token: string): Record<string, unknown> | null {
  const parts = token.split(".");
  if (parts.length < 3) {
    return null;
  }
  try {
    const json = Buffer.from(
      parts[1].replace(/-/g, "+").replace(/_/g, "/"),
      "base64",
    ).toString("utf8");
    return JSON.parse(json) as Record<string, unknown>;
  } catch {
    return null;
  }
}

function buildResult(accessToken: string, expiresOn: Date | null): MsalTokenResult {
  const claims = decodeJwtClaims(accessToken) ?? {};
  const oid = typeof claims.oid === "string" ? claims.oid : null;
  const tid = typeof claims.tid === "string" ? claims.tid : null;
  const exp = typeof claims.exp === "number" ? claims.exp * 1000 : null;
  const expiresAtUtc =
    expiresOn ?? (exp ? new Date(exp) : new Date(Date.now() + 3_600_000));
  return { token: accessToken, expiresAtUtc, oid, tid };
}

function isSubstrateToken(result: MsalTokenResult): boolean {
  const claims = decodeJwtClaims(result.token);
  const aud = claims && typeof claims.aud === "string" ? claims.aud : "";
  return aud.includes("substrate.office.com");
}

/**
 * Acquires a Substrate access token through MSAL. Prefers a silent
 * refresh-token exchange from the persisted cache (no browser); only when that
 * is unavailable and `allowInteractive` is set does it fall back to a headless
 * authorization-code + PKCE capture that reuses the saved browser sign-in
 * cookies. Returns null on any failure so the caller can fall back to the
 * legacy Playwright cookie path.
 */
export async function acquireSubstrateToken(
  options: MsalAcquireOptions = {},
): Promise<MsalTokenResult | null> {
  const cachePath = await getMsalCachePath();
  const app = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY },
  });
  if (await fileExists(cachePath)) {
    try {
      app.getTokenCache().deserialize(await fs.readFile(cachePath, "utf8"));
    } catch {
      // Corrupt cache; a fresh interactive flow will reseed it.
    }
  }

  const persistCache = async (): Promise<void> => {
    try {
      await fs.writeFile(cachePath, app.getTokenCache().serialize(), "utf8");
    } catch {
      // Non-fatal: cache persistence is best-effort.
    }
  };

  const silent = await acquireSilently(app, options.quiet);
  if (silent) {
    await persistCache();
    return isSubstrateToken(silent) ? silent : null;
  }

  if (options.allowInteractive !== false) {
    const interactive = await acquireInteractively(app, options);
    if (interactive) {
      await persistCache();
      return isSubstrateToken(interactive) ? interactive : null;
    }
  }

  return null;
}

async function acquireSilently(
  app: msal.PublicClientApplication,
  quiet?: boolean,
): Promise<MsalTokenResult | null> {
  let accounts: msal.AccountInfo[];
  try {
    accounts = await app.getTokenCache().getAllAccounts();
  } catch {
    return null;
  }
  if (accounts.length === 0) {
    return null;
  }
  try {
    const result = await app.acquireTokenSilent({
      scopes: SCOPES,
      account: accounts[0],
    });
    return buildResult(result.accessToken, result.expiresOn ?? null);
  } catch (error) {
    if (!quiet) {
      console.error(
        `[msal] Silent token refresh failed: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
    return null;
  }
}

async function acquireInteractively(
  app: msal.PublicClientApplication,
  options: MsalAcquireOptions,
): Promise<MsalTokenResult | null> {
  let chromium: typeof import("playwright").chromium;
  try {
    ({ chromium } = await import("playwright"));
  } catch {
    return null;
  }

  const cryptoProvider = new msal.CryptoProvider();
  const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
  const authUrl = await app.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
  });

  const statePath =
    options.browserStatePath ?? (await getBrowserStatePath());
  const storageState = await loadSanitizedStorageState(statePath);
  const headless = options.headless ?? true;
  const launchArgs = [
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-search-engine-choice-screen",
  ];
  const channel = resolveBrowserChannel(options.browser);

  let browser: import("playwright").Browser;
  try {
    browser = channel
      ? await chromium.launch({ headless, channel, args: launchArgs })
      : await chromium.launch({ headless, args: launchArgs });
  } catch {
    try {
      browser = await chromium.launch({ headless, args: launchArgs });
    } catch {
      return null;
    }
  }

  const context = await browser.newContext(
    storageState ? { storageState } : {},
  );
  const page = await context.newPage();
  let resolveCode: (code: string) => void = () => {};
  const codePromise = new Promise<string>((resolve) => {
    resolveCode = resolve;
  });
  page.on("request", (req) => {
    const url = req.url();
    if (url.includes("/oauth2/nativeclient") && url.includes("code=")) {
      const code = new URL(url).searchParams.get("code");
      if (code) {
        resolveCode(code);
      }
    }
  });

  try {
    await page.goto(authUrl, { waitUntil: "domcontentloaded" });
    const timeoutMs = headless ? 45_000 : 180_000;
    const code = await Promise.race([
      codePromise,
      new Promise<never>((_, reject) =>
        setTimeout(
          () => reject(new Error("timed out waiting for authorization code")),
          timeoutMs,
        ),
      ),
    ]);
    const result = await app.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
      codeVerifier: verifier,
    });
    return buildResult(result.accessToken, result.expiresOn ?? null);
  } catch (error) {
    if (!options.quiet) {
      console.error(
        `[msal] Interactive token acquisition failed: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
    return null;
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

function resolveBrowserChannel(
  browser: PlaywrightBrowserName | undefined,
): string | undefined {
  switch (browser) {
    case "edge":
      return "msedge";
    case "chrome":
      return "chrome";
    default:
      return undefined;
  }
}

type PlaywrightStorageState = Exclude<
  NonNullable<BrowserContextOptions["storageState"]>,
  string
>;
type PlaywrightCookie = PlaywrightStorageState["cookies"][number];

async function loadSanitizedStorageState(
  statePath: string,
): Promise<PlaywrightStorageState | null> {
  if (!(await fileExists(statePath))) {
    return null;
  }
  try {
    const state = JSON.parse(await fs.readFile(statePath, "utf8")) as {
      cookies?: PlaywrightCookie[];
    };
    if (!Array.isArray(state.cookies)) {
      return null;
    }
    // Keep only Microsoft identity cookies so the authorize request stays under
    // the header-size limit while still enabling silent SSO.
    const identityPattern =
      /(^|\.)(login\.microsoftonline\.com|login\.microsoft\.com|login\.live\.com|microsoftonline\.com|microsoftazuread-sso\.com)$/i;
    const cookies = state.cookies.filter((cookie) =>
      identityPattern.test(String(cookie?.domain ?? "").replace(/^\./, "")),
    );
    return { cookies, origins: [] };
  } catch {
    return null;
  }
}
