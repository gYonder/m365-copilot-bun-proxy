import { fetchTokenWithPlaywright } from "../cli/playwright-token";
import {
  acquireSubstrateToken,
  isValidSubstrateToken,
  type MsalTokenResult,
} from "../cli/msal-auth";
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

type TokenProviderDependencies = {
  getTokenPath: typeof getTokenPath;
  getBrowserStatePath: typeof getBrowserStatePath;
  loadToken: typeof loadToken;
  saveToken: typeof saveToken;
  acquireSubstrateToken: typeof acquireSubstrateToken;
  fetchTokenWithPlaywright: typeof fetchTokenWithPlaywright;
};

const defaultDependencies: TokenProviderDependencies = {
  getTokenPath,
  getBrowserStatePath,
  loadToken,
  saveToken,
  acquireSubstrateToken,
  fetchTokenWithPlaywright,
};

export class ProxyTokenProvider {
  private readonly tokenPathPromise: Promise<string>;
  private readonly browserStatePathPromise: Promise<string>;
  private readonly ignoreIncomingAuthorizationHeader: boolean;
  private readonly playwrightBrowser: PlaywrightBrowser;
  private readonly msalAuthEnabled: boolean;
  private readonly dependencies: TokenProviderDependencies;
  private inFlightAcquirePromise: Promise<string | null> | null = null;

  constructor(options?: {
    ignoreIncomingAuthorizationHeader?: boolean;
    playwrightBrowser?: PlaywrightBrowser;
    msalAuthEnabled?: boolean;
    dependencies?: Partial<TokenProviderDependencies>;
  }) {
    this.dependencies = { ...defaultDependencies, ...options?.dependencies };
    this.tokenPathPromise = this.dependencies.getTokenPath();
    this.browserStatePathPromise = this.dependencies.getBrowserStatePath();
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
    const tokenState = await this.dependencies.loadToken(tokenPath);
    if (!isTokenStateValid(tokenState) || !isValidPersistedSubstrateToken(tokenState)) {
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
      await this.dependencies.fetchTokenWithPlaywright(tokenPath, browserStatePath, {
        quiet: true,
        browser: this.playwrightBrowser,
      });
    } catch {
      return null;
    }

    const fetched = await this.dependencies.loadToken(tokenPath);
    return isTokenStateValid(fetched) && isValidPersistedSubstrateToken(fetched)
      ? `Bearer ${fetched.token}`
      : null;
  }

  private async tryAcquireViaMsal(
    tokenPath: string,
    browserStatePath: string,
  ): Promise<string | null> {
    try {
      const result = await this.dependencies.acquireSubstrateToken({
        tokenPath,
        browserStatePath,
        browser: this.playwrightBrowser,
        allowInteractive: true,
        headless: true,
        quiet: true,
      });
      if (!result?.token?.trim() || !isValidSubstrateToken(result)) {
        return null;
      }
      await this.dependencies.saveToken(tokenPath, result.token, result.expiresAtUtc, {
        oid: result.oid,
        tid: result.tid,
      });
      return `Bearer ${result.token}`;
    } catch {
      return null;
    }
  }
}

function isValidPersistedSubstrateToken(tokenState: TokenState): boolean {
  const result: MsalTokenResult = {
    token: tokenState.token,
    expiresAtUtc: new Date(tokenState.expiresAtUtc),
    oid: tokenState.oid ?? null,
    tid: tokenState.tid ?? null,
  };
  return isValidSubstrateToken(result, tokenState.tid);
}

function isTokenStateValid(tokenState: TokenState | null): tokenState is TokenState {
  if (!tokenState?.token?.trim()) {
    return false;
  }
  const expiresAtUtc = new Date(tokenState.expiresAtUtc);
  return expiresAtUtc.getTime() > Date.now() + TOKEN_EXPIRY_SKEW_MS;
}
