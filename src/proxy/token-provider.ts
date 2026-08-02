import { fetchTokenWithPlaywright } from "../cli/playwright-token";
import { acquireSubstrateToken } from "../cli/msal-auth";
import {
  getBrowserStatePath,
  getTokenPath,
  loadToken,
  saveToken,
  type TokenState,
} from "../cli/token-helpers";
import { PlaywrightBrowsers, type PlaywrightBrowser } from "./types";
import { normalizeBearerToken } from "./utils";

const TOKEN_EXPIRY_SKEW_MS = 60_000;

export class ProxyTokenProvider {
  private readonly tokenPathPromise: Promise<string>;
  private readonly browserStatePathPromise: Promise<string>;
  private readonly ignoreIncomingAuthorizationHeader: boolean;
  private readonly playwrightBrowser: PlaywrightBrowser;
  private readonly msalAuthEnabled: boolean;
  private inFlightAcquirePromise: Promise<string | null> | null = null;

  constructor(options?: {
    ignoreIncomingAuthorizationHeader?: boolean;
    playwrightBrowser?: PlaywrightBrowser;
    msalAuthEnabled?: boolean;
  }) {
    this.tokenPathPromise = getTokenPath();
    this.browserStatePathPromise = getBrowserStatePath();
    this.ignoreIncomingAuthorizationHeader =
      options?.ignoreIncomingAuthorizationHeader ?? true;
    this.playwrightBrowser =
      options?.playwrightBrowser ?? PlaywrightBrowsers.Edge;
    this.msalAuthEnabled = options?.msalAuthEnabled ?? true;
  }

  async resolveAuthorizationHeader(
    rawAuthorizationHeader: string | null | undefined,
  ): Promise<string | null> {
    if (!this.ignoreIncomingAuthorizationHeader) {
      const providedHeader = normalizeBearerToken(rawAuthorizationHeader);
      if (providedHeader) {
        return providedHeader;
      }
    }

    const cachedHeader = await this.tryGetCachedAuthorizationHeader();
    if (cachedHeader) {
      return cachedHeader;
    }

    return this.acquireAuthorizationHeader();
  }

  private async tryGetCachedAuthorizationHeader(): Promise<string | null> {
    const tokenPath = await this.tokenPathPromise;
    const tokenState = await loadToken(tokenPath);
    if (!isTokenStateValid(tokenState)) {
      return null;
    }
    return `Bearer ${tokenState.token}`;
  }

  private async acquireAuthorizationHeader(): Promise<string | null> {
    const pendingAcquire =
      this.inFlightAcquirePromise ?? this.acquireFreshAuthorizationHeader();
    if (!this.inFlightAcquirePromise) {
      this.inFlightAcquirePromise = pendingAcquire;
    }
    try {
      return await pendingAcquire;
    } finally {
      if (this.inFlightAcquirePromise === pendingAcquire) {
        this.inFlightAcquirePromise = null;
      }
    }
  }

  private async acquireFreshAuthorizationHeader(): Promise<string | null> {
    const [tokenPath, browserStatePath] = await Promise.all([
      this.tokenPathPromise,
      this.browserStatePathPromise,
    ]);

    // Primary path: MSAL OAuth2 auth-code + PKCE. A silent refresh-token
    // exchange renews the substrate token without any browser or cookie replay,
    // which is what previously overflowed request headers on renewal.
    if (this.msalAuthEnabled) {
      const msalHeader = await this.tryAcquireViaMsal(
        tokenPath,
        browserStatePath,
      );
      if (msalHeader) {
        return msalHeader;
      }
    }

    // Fallback: legacy Playwright cookie capture.
    try {
      await fetchTokenWithPlaywright(tokenPath, browserStatePath, {
        quiet: true,
        browser: this.playwrightBrowser,
      });
    } catch {
      return null;
    }

    const fetched = await loadToken(tokenPath);
    return isTokenStateValid(fetched) ? `Bearer ${fetched.token}` : null;
  }

  private async tryAcquireViaMsal(
    tokenPath: string,
    browserStatePath: string,
  ): Promise<string | null> {
    try {
      const result = await acquireSubstrateToken({
        browserStatePath,
        browser: this.playwrightBrowser,
        allowInteractive: true,
        headless: true,
        quiet: true,
      });
      if (!result?.token?.trim()) {
        return null;
      }
      await saveToken(tokenPath, result.token, result.expiresAtUtc, {
        oid: result.oid,
        tid: result.tid,
      });
      return `Bearer ${result.token}`;
    } catch {
      return null;
    }
  }
}

function isTokenStateValid(tokenState: TokenState | null): tokenState is TokenState {
  if (!tokenState?.token?.trim()) {
    return false;
  }
  const expiresAtUtc = new Date(tokenState.expiresAtUtc);
  return expiresAtUtc.getTime() > Date.now() + TOKEN_EXPIRY_SKEW_MS;
}
