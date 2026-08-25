import { promises as fs } from "node:fs";
import path from "node:path";
import * as msal from "@azure/msal-node";
import {
  PersistenceCachePlugin,
  PersistenceCreator,
} from "@azure/msal-node-extensions";
import type { BrowserContextOptions } from "playwright";
import {
  getBrowserProfilePath,
  getBrowserStatePath,
  getTokenPath,
  loadToken,
} from "./token-helpers";
import type { PlaywrightBrowserName } from "./playwright-token";

// First-party public-client app id that M365 Copilot ("Sydney") web uses. The
// authorization-code + PKCE flow against it yields substrate access tokens
// directly, so token renewal becomes a silent refresh-token call instead of a
// full browser cookie replay (which is what overflowed request headers).
const CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const AUTHORITY = "https://login.microsoftonline.com/common";
const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";
const RESOURCE_SCOPES = [
  "https://substrate.office.com/sydney/M365Chat.Read",
  "https://substrate.office.com/sydney/sydney.readwrite",
];
const SCOPES = ["openid", "profile", "offline_access", ...RESOURCE_SCOPES];
const DESIGNER_SCOPES = ["https://designer.microsoft.com/.default"];
const DESIGNER_AUDIENCES = new Set([
  "https://designer.microsoft.com",
  "designer.microsoft.com",
]);
const TOKEN_EXPIRY_SKEW_MS = 60_000;
const SUBSTRATE_AUDIENCE = "https://substrate.office.com";
const SYDNEY_AUDIENCE = "https://substrate.office.com/sydney";

export type MsalTokenResult = {
  token: string;
  expiresAtUtc: Date;
  oid: string | null;
  tid: string | null;
};

export type MsalAcquireOptions = {
  cachePath?: string;
  tokenPath?: string;
  browserStatePath?: string;
  browser?: PlaywrightBrowserName;
  allowInteractive?: boolean;
  headless?: boolean;
  quiet?: boolean;
  appFactory?: () => MsalPublicClient;
};

export type MsalAuthStatus =
  | "silent_success"
  | "interactive_required"
  | "interactive_success"
  | "cache_unavailable"
  | "account_unavailable"
  | "invalid_token"
  | "interaction_failed";

export type MsalAcquireResult = {
  token: MsalTokenResult | null;
  status: MsalAuthStatus;
  message?: string;
};

export type DesignerAcquireOptions = Omit<MsalAcquireOptions, "tokenPath">;

type MsalTokenCache = {
  deserialize(value: string): void;
  serialize(): string;
  getAllAccounts(): Promise<msal.AccountInfo[]>;
};

type MsalPublicClient = Pick<
  msal.PublicClientApplication,
  "acquireTokenSilent" | "acquireTokenByCode" | "getAuthCodeUrl"
> & {
  getTokenCache(): MsalTokenCache;
};

async function getMsalCachePath(): Promise<string> {
  const tokenPath = await getTokenPath();
  return path.join(path.dirname(tokenPath), "msal-cache.json");
}

const CACHE_SERVICE = "com.gyonder.m365-copilot-bun-proxy";

async function createMsalApp(
  cachePath: string,
  cacheName: "sydney" | "designer",
  appFactory?: () => MsalPublicClient,
): Promise<{ app: MsalPublicClient; manualCache: boolean; legacyCache: string | null }> {
  if (appFactory) {
    return { app: appFactory(), manualCache: true, legacyCache: null };
  }

  // The extension stores the MSAL cache in the macOS Keychain and leaves only
  // a harmless marker at cachePath. Read an older plaintext cache before the
  // extension gets a chance to replace that marker so it can be migrated once.
  let legacyCache: string | null = null;
  try {
    const old = await fs.readFile(cachePath, "utf8");
    if (old.trim() && old.trim() !== "{}") legacyCache = old;
  } catch {
    // First run.
  }

  const persistence = await PersistenceCreator.createPersistence({
    cachePath,
    serviceName: `${CACHE_SERVICE}.${cacheName}`,
    accountName: "default",
  });
  const app = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY },
    cache: { cachePlugin: new PersistenceCachePlugin(persistence) },
  });
  if (legacyCache) {
    try {
      app.getTokenCache().deserialize(legacyCache);
    } catch {
      legacyCache = null;
    }
  }
  return { app, manualCache: false, legacyCache };
}

async function writeOwnerOnlyAtomic(filePath: string, value: string): Promise<void> {
  const directory = path.dirname(filePath);
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const temporaryPath = path.join(
    directory,
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.tmp`,
  );
  try {
    await fs.writeFile(temporaryPath, value, { encoding: "utf8", mode: 0o600 });
    await fs.chmod(temporaryPath, 0o600);
    await fs.rename(temporaryPath, filePath);
    await fs.chmod(filePath, 0o600);
  } catch (error) {
    await fs.unlink(temporaryPath).catch(() => {});
    throw error;
  }
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

function buildResult(accessToken: string, expiresOn: Date | null): MsalTokenResult | null {
  const claims = decodeJwtClaims(accessToken);
  if (!claims) return null;
  const oid = typeof claims.oid === "string" ? claims.oid : null;
  const tid = typeof claims.tid === "string" ? claims.tid : null;
  const exp = typeof claims.exp === "number" ? claims.exp * 1000 : null;
  const expiresAtUtc = exp ? new Date(exp) : expiresOn;
  if (!expiresAtUtc || expiresAtUtc.getTime() <= Date.now() + TOKEN_EXPIRY_SKEW_MS) {
    return null;
  }
  return { token: accessToken, expiresAtUtc, oid, tid };
}

export function isValidSubstrateToken(
  result: MsalTokenResult,
  requiredTenantId?: string | null,
): boolean {
  const claims = decodeJwtClaims(result.token);
  if (!claims) return false;
  const aud = typeof claims.aud === "string" ? claims.aud.toLowerCase() : "";
  const tid = typeof claims.tid === "string" ? claims.tid : "";
  const exp = typeof claims.exp === "number" ? claims.exp * 1000 : 0;
  const scopeClaim = typeof claims.scp === "string" ? claims.scp : "";
  const scopes = new Set(scopeClaim.split(/\s+/).filter(Boolean));
  const claimOid = typeof claims.oid === "string" ? claims.oid : "";
  return (
    aud === SUBSTRATE_AUDIENCE ||
    aud === SYDNEY_AUDIENCE ||
    aud === "substrate.office.com"
  ) &&
    exp > Date.now() + TOKEN_EXPIRY_SKEW_MS &&
    !!tid &&
    (!requiredTenantId || tid === requiredTenantId) &&
    (!result.tid || result.tid === tid) &&
    (!result.oid || result.oid === claimOid) &&
    RESOURCE_SCOPES.every((scope) => scopes.has(scope.split("/").at(-1) ?? scope));
}

export function isValidDesignerToken(result: MsalTokenResult): boolean {
  const claims = decodeJwtClaims(result.token);
  if (!claims) return false;
  const audience =
    typeof claims.aud === "string"
      ? claims.aud.toLowerCase().replace(/\/$/, "")
      : "";
  const expiry = typeof claims.exp === "number" ? claims.exp * 1000 : 0;
  return (
    DESIGNER_AUDIENCES.has(audience) &&
    expiry > Date.now() + TOKEN_EXPIRY_SKEW_MS
  );
}

export async function acquireDesignerToken(
  options: DesignerAcquireOptions = {},
): Promise<MsalTokenResult | null> {
  const substrateCachePath = await getMsalCachePath();
  const cachePath =
    options.cachePath ??
    path.join(path.dirname(substrateCachePath), "msal-designer-cache.json");
  let created: Awaited<ReturnType<typeof createMsalApp>>;
  try {
    created = await createMsalApp(cachePath, "designer", options.appFactory);
  } catch {
    return null;
  }
  const { app, manualCache } = created;

  if (manualCache && (await fileExists(cachePath))) {
    try {
      await fs.chmod(cachePath, 0o600);
      app.getTokenCache().deserialize(await fs.readFile(cachePath, "utf8"));
    } catch {
      // The Designer cache is isolated from the Sydney cache and can be reseeded.
    }
  }

  let accounts: msal.AccountInfo[];
  try {
    accounts = await app.getTokenCache().getAllAccounts();
  } catch {
    return null;
  }
  if (accounts.length !== 1) return null;

  try {
    const acquired = await app.acquireTokenSilent({
      scopes: DESIGNER_SCOPES,
      account: accounts[0],
    });
    const result = buildResult(acquired.accessToken, acquired.expiresOn ?? null);
    if (!result || !isValidDesignerToken(result)) return null;
    if (manualCache) {
      await writeOwnerOnlyAtomic(cachePath, app.getTokenCache().serialize());
    }
    return result;
  } catch {
    return null;
  }
}

const MSAL_CACHE_LOCK_RETRY_MS = 100;
const MSAL_CACHE_LOCK_TIMEOUT_MS = 8_000;
const MSAL_CACHE_LOCK_STALE_MS = 30_000;

/**
 * Serializes silent refresh-token redemptions across processes. Entra rotates
 * refresh tokens single-use on every redemption, and the CLI token login/fetch
 * process plus the long-running proxy share one MSAL cache file: concurrent
 * redemptions invalidate the loser's refresh token and auth stays broken until
 * an interactive sign-in. Best-effort advisory lock; on timeout the operation
 * proceeds unlocked rather than blocking authentication entirely.
 */
export async function withMsalCacheLock<T>(
  cachePath: string,
  operation: () => Promise<T>,
): Promise<T> {
  const lockPath = `${cachePath}.lock`;
  const deadline = Date.now() + MSAL_CACHE_LOCK_TIMEOUT_MS;
  let acquired = false;
  while (!acquired) {
    try {
      const handle = await fs.open(lockPath, "wx");
      await handle.writeFile(`${process.pid} ${new Date().toISOString()}\n`);
      await handle.close();
      acquired = true;
    } catch (error) {
      if ((error as NodeJS.ErrnoException)?.code !== "EEXIST") throw error;
      const lockAgeMs = await fs
        .stat(lockPath)
        .then((stats) => Date.now() - stats.mtimeMs, () => 0);
      if (lockAgeMs > MSAL_CACHE_LOCK_STALE_MS) {
        // Atomic steal: rename succeeds for exactly one contender, so a stale
        // lock left by a crashed owner is reclaimed without a thundering herd.
        const stolenPath = `${lockPath}.${process.pid}.stale`;
        await fs
          .rename(lockPath, stolenPath)
          .then(() => fs.rm(stolenPath, { force: true }))
          .catch(() => undefined);
        continue;
      }
      if (Date.now() >= deadline) break;
      await new Promise((resolve) => setTimeout(resolve, MSAL_CACHE_LOCK_RETRY_MS));
    }
  }
  try {
    return await operation();
  } finally {
    if (acquired) {
      await fs.rm(lockPath, { force: true }).catch(() => undefined);
    }
  }
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
  const result = await acquireSubstrateTokenDetailed(options);
  return result.token;
}

export async function acquireSubstrateTokenDetailed(
  options: MsalAcquireOptions = {},
): Promise<MsalAcquireResult> {
  const cachePath = options.cachePath ?? (await getMsalCachePath());
  const tokenPath = options.tokenPath ?? (await getTokenPath());
  const priorToken = await loadToken(tokenPath);
  let created: Awaited<ReturnType<typeof createMsalApp>>;
  try {
    created = await createMsalApp(cachePath, "sydney", options.appFactory);
  } catch {
    return { token: null, status: "cache_unavailable" };
  }
  const { app, manualCache } = created;

  const persistCache = async (): Promise<boolean> => {
    if (!manualCache) return true;
    try {
      await writeOwnerOnlyAtomic(cachePath, app.getTokenCache().serialize());
      return true;
    } catch {
      return false;
    }
  };

  // The silent path reads the shared cache, redeems the single-use refresh
  // token, and writes the rotated state back. That sequence must not interleave
  // with another process doing the same thing.
  const attemptSilent = async (): Promise<MsalAcquireResult | null> => {
    if (manualCache && (await fileExists(cachePath))) {
      try {
        await fs.chmod(cachePath, 0o600);
        app.getTokenCache().deserialize(await fs.readFile(cachePath, "utf8"));
      } catch {
        // Corrupt cache; a fresh interactive flow will reseed it.
      }
    }
    const silent = await acquireSilently(app, priorToken, options.quiet);
    if (!silent) return null;
    if (!(await persistCache())) return { token: null, status: "cache_unavailable" };
    return isValidSubstrateToken(silent, priorToken?.tid)
      ? { token: silent, status: "silent_success" }
      : { token: null, status: "invalid_token" };
  };

  const silentResult = manualCache
    ? await withMsalCacheLock(cachePath, attemptSilent)
    : await attemptSilent();
  if (silentResult) return silentResult;

  if (options.allowInteractive !== false) {
    const interactive = await acquireInteractively(app, options);
    if (interactive) {
      if (!(await persistCache())) return { token: null, status: "cache_unavailable" };
      return isValidSubstrateToken(interactive)
        ? { token: interactive, status: "interactive_success" }
        : { token: null, status: "invalid_token" };
    }
    return { token: null, status: "interaction_failed" };
  }

  return {
    token: null,
    status: (await app.getTokenCache().getAllAccounts().catch(() => [])).length
      ? "interactive_required"
      : "account_unavailable",
  };
}

async function acquireSilently(
  app: MsalPublicClient,
  priorToken: Awaited<ReturnType<typeof loadToken>>,
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
  let account: msal.AccountInfo | null = null;
  if (priorToken?.tid || priorToken?.oid) {
    const matching = accounts.filter((candidate) =>
      (!priorToken.tid || candidate.tenantId === priorToken.tid) &&
      (!priorToken.oid || candidate.localAccountId === priorToken.oid),
    );
    if (matching.length === 1) account = matching[0] ?? null;
  }
  if (!account && accounts.length === 1) account = accounts[0] ?? null;
  if (!account) return null;
  try {
    const result = await app.acquireTokenSilent({
      scopes: SCOPES,
      account,
    });
    return buildResult(result.accessToken, result.expiresOn ?? null);
  } catch {
    if (!quiet) {
      console.error("[msal] Silent token refresh failed; interactive sign-in may be required.");
    }
    return null;
  }
}

async function acquireInteractively(
  app: MsalPublicClient,
  options: MsalAcquireOptions,
): Promise<MsalTokenResult | null> {
  let chromium: typeof import("playwright").chromium;
  try {
    ({ chromium } = await import("playwright"));
  } catch {
    return null;
  }

  let verifier: string;
  let authUrl: string;
  try {
    const cryptoProvider = new msal.CryptoProvider();
    const generated = await cryptoProvider.generatePkceCodes();
    verifier = generated.verifier;
    authUrl = await app.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
      codeChallenge: generated.challenge,
      codeChallengeMethod: "S256",
    });
  } catch {
    return null;
  }

  const statePath =
    options.browserStatePath ?? (await getBrowserStatePath());
  const profilePath = await getBrowserProfilePath();
  const storageState = await loadSanitizedStorageState(statePath);
  const headless = options.headless ?? false;
  const launchArgs = [
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-search-engine-choice-screen",
  ];
  const channel = resolveBrowserChannel(options.browser);

  let context: import("playwright").BrowserContext;
  try {
    context = await chromium.launchPersistentContext(profilePath, {
      headless,
      ...(channel ? { channel } : {}),
      args: launchArgs,
    });
  } catch {
    try {
      context = await chromium.launchPersistentContext(profilePath, {
        headless,
        args: launchArgs,
      });
    } catch {
      return null;
    }
  }
  if (storageState?.cookies?.length) {
    await context.addCookies(storageState.cookies).catch(() => {});
  }
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
      console.error("[msal] Interactive token acquisition did not complete.");
    }
    return null;
  } finally {
    await context.close().catch(() => {});
    await context.close().catch(() => {});
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
