import { fetchTokenWithPlaywright } from "../cli/playwright-token";
import {
  acquireDesignerToken,
  acquireSubstrateToken,
  isValidDesignerToken,
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
import type { BridgeObservability } from "./observability";

const TOKEN_EXPIRY_SKEW_MS = 60_000;

type TokenProviderDependencies = {
  getTokenPath: typeof getTokenPath;
  getBrowserStatePath: typeof getBrowserStatePath;
  loadToken: typeof loadToken;
  saveToken: typeof saveToken;
  acquireSubstrateToken: typeof acquireSubstrateToken;
  acquireDesignerToken: typeof acquireDesignerToken;
  fetchTokenWithPlaywright: typeof fetchTokenWithPlaywright;
};

const defaultDependencies: TokenProviderDependencies = {
  getTokenPath,
  getBrowserStatePath,
  loadToken,
  saveToken,
  acquireSubstrateToken,
  acquireDesignerToken,
  fetchTokenWithPlaywright,
};

export class ProxyTokenProvider {
  private readonly tokenPathPromise: Promise<string>;
  private readonly browserStatePathPromise: Promise<string>;
  private readonly ignoreIncomingAuthorizationHeader: boolean;
  private readonly playwrightBrowser: PlaywrightBrowser;
  private readonly msalAuthEnabled: boolean;
  private readonly dependencies: TokenProviderDependencies;
  private readonly observability: BridgeObservability | null;
  private inFlightAcquirePromise: Promise<string | null> | null = null;
  private inFlightDesignerAcquirePromise: Promise<string | null> | null = null;
  private authorizationHeader: string | null = null;
  private authorizationExpiresAtMs = 0;
  private designerAuthorizationHeader: string | null = null;
  private designerExpiresAtMs = 0;

  constructor(options?: {
    ignoreIncomingAuthorizationHeader?: boolean;
    playwrightBrowser?: PlaywrightBrowser;
    msalAuthEnabled?: boolean;
    dependencies?: Partial<TokenProviderDependencies>;
    observability?: BridgeObservability;
  }) {
    this.dependencies = { ...defaultDependencies, ...options?.dependencies };
    this.tokenPathPromise = this.dependencies.getTokenPath();
    this.browserStatePathPromise = this.dependencies.getBrowserStatePath();
    this.ignoreIncomingAuthorizationHeader =
      options?.ignoreIncomingAuthorizationHeader ?? true;
    this.playwrightBrowser =
      options?.playwrightBrowser ?? PlaywrightBrowsers.Edge;
    this.msalAuthEnabled = options?.msalAuthEnabled ?? true;
    this.observability = options?.observability ?? null;
  }

  async resolveAuthorizationHeader(
    rawAuthorizationHeader: string | null | undefined,
  ): Promise<string | null> {
    if (!this.ignoreIncomingAuthorizationHeader) {
      const providedHeader = normalizeBearerToken(rawAuthorizationHeader);
      if (providedHeader) {
        this.recordAuthPath("incoming", true);
        return providedHeader;
      }
    }

    if (
      this.authorizationHeader &&
      this.authorizationExpiresAtMs > Date.now() + TOKEN_EXPIRY_SKEW_MS
    ) {
      this.recordAuthPath("memory_cache", true);
      return this.authorizationHeader;
    }

    const cachedHeader = await this.tryGetCachedAuthorizationHeader();
    if (cachedHeader) {
      this.recordAuthPath("persisted_cache", true);
      return cachedHeader;
    }

    return this.acquireAuthorizationHeader();
  }

  async resolveDesignerAuthorizationHeader(): Promise<string | null> {
    if (
      this.designerAuthorizationHeader &&
      this.designerExpiresAtMs > Date.now() + TOKEN_EXPIRY_SKEW_MS
    ) {
      this.recordAuthPath("designer_memory_cache", true, "designer");
      return this.designerAuthorizationHeader;
    }
    const pending =
      this.inFlightDesignerAcquirePromise ??
      this.acquireDesignerAuthorizationHeader();
    if (!this.inFlightDesignerAcquirePromise) {
      this.inFlightDesignerAcquirePromise = pending;
    }
    try {
      return await pending;
    } finally {
      if (this.inFlightDesignerAcquirePromise === pending) {
        this.inFlightDesignerAcquirePromise = null;
      }
    }
  }

  private async acquireDesignerAuthorizationHeader(): Promise<string | null> {
    if (!this.msalAuthEnabled) {
      this.recordAuthPath("designer_msal_disabled", false, "designer");
      return null;
    }
    try {
      const result = await this.dependencies.acquireDesignerToken({
        browserStatePath: await this.browserStatePathPromise,
        browser: this.playwrightBrowser,
        allowInteractive: false,
        headless: true,
        quiet: true,
      });
      if (!result || !isValidDesignerToken(result)) {
        this.recordAuthPath("designer_msal", false, "designer");
        return null;
      }
      this.designerAuthorizationHeader = "Bearer " + result.token;
      this.designerExpiresAtMs = result.expiresAtUtc.getTime();
      this.recordAuthPath("designer_msal", true, "designer");
      return this.designerAuthorizationHeader;
    } catch {
      this.recordAuthPath("designer_msal", false, "designer");
      return null;
    }
  }

  private async tryGetCachedAuthorizationHeader(): Promise<string | null> {
    const tokenPath = await this.tokenPathPromise;
    const tokenState = await this.dependencies.loadToken(tokenPath);
    if (!isTokenStateValid(tokenState) || !isValidPersistedSubstrateToken(tokenState)) {
      return null;
    }
    return this.rememberAuthorizationHeader(tokenState);
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

    // The proxy is a background service. Only MSAL's persisted refresh-token
    // path is allowed here; interactive browser work belongs to the explicit
    // CLI/launcher login command and must never happen during an API request.
    if (this.msalAuthEnabled) {
      const msalHeader = await this.tryAcquireViaMsal(
        tokenPath,
        browserStatePath,
      );
      if (msalHeader) {
        this.recordAuthPath("msal", true);
        return msalHeader;
      }
      this.recordAuthPath("msal", false);
    }
    this.recordAuthPath("interactive_required", false);
    return null;
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
        allowInteractive: false,
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
      return this.rememberAuthorizationHeader({
        token: result.token,
        expiresAtUtc: result.expiresAtUtc.toISOString(),
        oid: result.oid ?? undefined,
        tid: result.tid ?? undefined,
      });
    } catch {
      return null;
    }
  }

  private rememberAuthorizationHeader(tokenState: TokenState): string {
    this.authorizationHeader = `Bearer ${tokenState.token}`;
    this.authorizationExpiresAtMs = new Date(tokenState.expiresAtUtc).getTime();
    return this.authorizationHeader;
  }

  private recordAuthPath(
    path: string,
    success: boolean,
    audience: "sydney" | "designer" = "sydney",
  ): void {
    this.observability?.record("auth_path", { path, success, audience });
  }
}

function isValidPersistedSubstrateToken(tokenState: TokenState): boolean {
  const result: MsalTokenResult = {
    token: tokenState.token,
    expiresAtUtc: new Date(tokenState.expiresAtUtc),
    oid: tokenState.oid ?? null,
    tid: tokenState.tid ?? null,
  };
  return (
    Boolean(tokenState.tid?.trim()) &&
    isValidSubstrateToken(result, tokenState.tid)
  );
}

function isTokenStateValid(tokenState: TokenState | null): tokenState is TokenState {
  if (!tokenState?.token?.trim()) {
    return false;
  }
  const expiresAtUtc = new Date(tokenState.expiresAtUtc);
  return expiresAtUtc.getTime() > Date.now() + TOKEN_EXPIRY_SKEW_MS;
}
