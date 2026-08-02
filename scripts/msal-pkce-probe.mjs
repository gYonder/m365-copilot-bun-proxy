// Feasibility probe: acquire a Microsoft 365 Copilot ("Sydney") substrate access
// token via the first-party public-client OAuth2 authorization-code + PKCE flow
// (the approach used by cramt/m365-copilot-proxy), instead of replaying the
// Playwright browser cookie jar. Proves whether MSAL can replace fragile cookie
// reuse for substrate auth.
//
// Usage (from the proxy dir):
//   bun scripts/msal-pkce-probe.mjs            # headless, silent SSO from saved session
//   bun scripts/msal-pkce-probe.mjs --headed   # visible browser for interactive sign-in
//   bun scripts/msal-pkce-probe.mjs --write-token   # also write token.json for an e2e proxy test
//
// It never prints the access token; only non-secret claims (aud/scp/tid/oid/exp).

import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import * as msal from "@azure/msal-node";
import { chromium } from "playwright";

const CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const AUTHORITY = "https://login.microsoftonline.com/common";
const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = [
  "https://substrate.office.com/sydney/M365Chat.Read",
  "https://substrate.office.com/sydney/sydney.readwrite",
];

const argv = new Set(process.argv.slice(2));
const HEADED = argv.has("--headed");
const WRITE_TOKEN = argv.has("--write-token");

function cliDataDir() {
  const base = process.env.LOCALAPPDATA ?? path.join(os.homedir(), ".local", "share");
  return path.join(base, "M365 Copilot Bun Proxy", "Cli");
}

async function fileExists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

function decodeJwtClaims(token) {
  const parts = token.split(".");
  if (parts.length < 2) return null;
  try {
    const json = Buffer.from(
      parts[1].replace(/-/g, "+").replace(/_/g, "/"),
      "base64",
    ).toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

async function loadSanitizedStorageState(statePath) {
  if (!(await fileExists(statePath))) return null;
  try {
    const state = JSON.parse(await fs.readFile(statePath, "utf8"));
    if (!state || !Array.isArray(state.cookies)) return null;
    // Keep only the small Microsoft identity/SSO cookies so the AAD authorize
    // request stays under the header limit and can silently issue a code.
    const idPattern =
      /(^|\.)(login\.microsoftonline\.com|login\.microsoft\.com|login\.live\.com|microsoftonline\.com|microsoftazuread-sso\.com)$/i;
    const cookies = state.cookies.filter((c) =>
      idPattern.test(String(c?.domain ?? "").replace(/^\./, "")),
    );
    return { cookies, origins: [] };
  } catch {
    return null;
  }
}

async function launchEdge() {
  const args = [
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-search-engine-choice-screen",
  ];
  try {
    return await chromium.launch({ headless: !HEADED, channel: "msedge", args });
  } catch {
    console.log("[probe] Edge channel unavailable, falling back to Chromium.");
    return chromium.launch({ headless: !HEADED, args });
  }
}

async function acquireViaBrowser(app) {
  const cryptoProvider = new msal.CryptoProvider();
  const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
  const authUrl = await app.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
  });

  const statePath = path.join(cliDataDir(), "browser-state.json");
  const storageState = await loadSanitizedStorageState(statePath);
  console.log(
    `[probe] Launching Edge (${HEADED ? "headed" : "headless"}) ${
      storageState
        ? `with ${storageState.cookies.length} saved sign-in cookies`
        : "with no saved session"
    }.`,
  );
  const browser = await launchEdge();
  const context = await browser.newContext(storageState ? { storageState } : {});
  const page = await context.newPage();

  let resolveCode;
  const codePromise = new Promise((res) => {
    resolveCode = res;
  });
  page.on("request", (req) => {
    const u = req.url();
    if (u.includes("/oauth2/nativeclient") && u.includes("code=")) {
      const c = new URL(u).searchParams.get("code");
      if (c) {
        console.log("[probe] Captured auth code from nativeclient redirect.");
        resolveCode(c);
      }
    }
  });

  try {
    await page.goto(authUrl, { waitUntil: "domcontentloaded" });
    const timeoutMs = HEADED ? 180_000 : 45_000;
    const code = await Promise.race([
      codePromise,
      new Promise((_, rej) =>
        setTimeout(
          () =>
            rej(
              new Error(
                `Timed out waiting for auth code (landed on: ${page.url().slice(0, 80)}). ` +
                  (HEADED
                    ? ""
                    : "Silent SSO did not complete; re-run with --headed to sign in interactively."),
              ),
            ),
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
    return result;
  } finally {
    await context.close();
    await browser.close();
  }
}

async function main() {
  const cacheFile = path.join(cliDataDir(), "msal-cache.json");
  const app = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY },
  });
  if (await fileExists(cacheFile)) {
    try {
      app.getTokenCache().deserialize(await fs.readFile(cacheFile, "utf8"));
    } catch {}
  }

  let result = null;

  // 1) Try silent refresh from a previously cached account (proves the refresh path).
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      result = await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      console.log(
        `[probe] Acquired token SILENTLY for cached account ${accounts[0].username}.`,
      );
    } catch (e) {
      console.log(
        `[probe] Silent acquisition failed (${e?.message ?? e}); falling back to browser.`,
      );
    }
  }

  // 2) Otherwise do the interactive/SSO-silent authorization-code + PKCE flow.
  if (!result) {
    result = await acquireViaBrowser(app);
    console.log(
      `[probe] Acquired token via authorization-code + PKCE for ${result.account?.username}.`,
    );
  }

  await fs.writeFile(cacheFile, app.getTokenCache().serialize(), "utf8");

  const claims = decodeJwtClaims(result.accessToken) ?? {};
  const scpFromClaims = typeof claims.scp === "string" ? claims.scp : null;
  const report = {
    isJwt: result.accessToken.split(".").length === 3,
    aud: claims.aud ?? null,
    scp: scpFromClaims,
    grantedScopes: result.scopes ?? null,
    tid: claims.tid ?? null,
    oid: claims.oid ? `${String(claims.oid).slice(0, 8)}…` : null,
    appid: claims.appid ?? null,
    exp: result.expiresOn ? result.expiresOn.toISOString() : null,
  };
  console.log("[probe] Token claims:", JSON.stringify(report, null, 2));

  const audOk =
    typeof report.aud === "string" && report.aud.includes("substrate.office.com");
  const scopeOk =
    (scpFromClaims ?? "").toLowerCase().includes("sydney") ||
    (result.scopes ?? []).some((s) => s.toLowerCase().includes("sydney"));
  console.log(
    `[probe] audience is substrate/sydney: ${audOk}; sydney scope present: ${scopeOk}`,
  );

  if (WRITE_TOKEN) {
    if (!claims.oid || !claims.tid) {
      throw new Error(
        "Token missing oid/tid claims; cannot write token.json for the e2e test.",
      );
    }
    const tokenPath = path.join(cliDataDir(), "token.json");
    if (await fileExists(tokenPath)) {
      await fs.copyFile(tokenPath, `${tokenPath}.probe-backup`);
      console.log(`[probe] Backed up existing token.json -> token.json.probe-backup`);
    }
    await fs.writeFile(
      tokenPath,
      JSON.stringify(
        {
          token: result.accessToken,
          expiresAtUtc: (
            result.expiresOn ?? new Date(Date.now() + 3_600_000)
          ).toISOString(),
          oid: claims.oid,
          tid: claims.tid,
        },
        null,
        2,
      ),
      "utf8",
    );
    console.log("[probe] Wrote MSAL token to token.json for an end-to-end proxy round-trip.");
  }

  if (report.isJwt && audOk && scopeOk) {
    console.log("PROBE_RESULT: SUCCESS");
    process.exit(0);
  }
  console.log("PROBE_RESULT: TOKEN_ACQUIRED_BUT_WRONG_AUDIENCE_OR_SCOPE");
  process.exit(2);
}

main().catch((e) => {
  console.error(`[probe] FAILED: ${e?.stack ?? e}`);
  console.log("PROBE_RESULT: FAILURE");
  process.exit(1);
});
