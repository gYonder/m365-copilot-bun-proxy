import { describe, expect, test } from "bun:test";
import { ProxyTokenProvider } from "../src/proxy/token-provider";

const encode = (value: unknown) => Buffer.from(JSON.stringify(value)).toString("base64url");
const token = (overrides: Record<string, unknown> = {}) =>
  `${encode({ alg: "none" })}.${encode({
    aud: "https://substrate.office.com",
    exp: Math.floor(Date.now() / 1000) + 3_600,
    tid: "tenant-a",
    oid: "object-a",
    scp: "M365Chat.Read sydney.readwrite",
    ...overrides,
  })}.signature`;

const state = (value: string) => ({
  token: value,
  expiresAtUtc: new Date(Date.now() + 3_600_000).toISOString(),
  tid: "tenant-a",
  oid: "object-a",
});

const designerToken = (overrides: Record<string, unknown> = {}) =>
  token({
    aud: "https://designer.microsoft.com",
    scp: "Designer.ReadWrite",
    ...overrides,
  });

describe("ProxyTokenProvider", () => {
  test("acquires and caches a separate Designer-audience token", async () => {
    const value = designerToken();
    let acquisitions = 0;
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => null,
      acquireDesignerToken: async () => {
        acquisitions += 1;
        return {
          token: value,
          expiresAtUtc: new Date(Date.now() + 3_600_000),
          oid: "object-a",
          tid: "tenant-a",
        };
      },
    }});
    expect(await provider.resolveDesignerAuthorizationHeader()).toBe(
      "Bearer " + value,
    );
    expect(await provider.resolveDesignerAuthorizationHeader()).toBe(
      "Bearer " + value,
    );
    expect(acquisitions).toBe(1);
  });

  test("rejects a Sydney token returned for the Designer audience", async () => {
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => null,
      acquireDesignerToken: async () => ({
        token: token(),
        expiresAtUtc: new Date(Date.now() + 3_600_000),
        oid: "object-a",
        tid: "tenant-a",
      }),
    }});
    expect(await provider.resolveDesignerAuthorizationHeader()).toBeNull();
  });

  test("uses a validated cached Sydney token", async () => {
    const value = token();
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => state(value),
    }});
    expect(await provider.resolveAuthorizationHeader(null)).toBe(`Bearer ${value}`);
  });

  test("accepts a cached token with the observed Sydney resource audience", async () => {
    const value = token({ aud: "https://substrate.office.com/sydney" });
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => state(value),
    }});
    expect(await provider.resolveAuthorizationHeader(null)).toBe("Bearer " + value);
  });

  test("rejects an invalid cached token and falls back to Playwright", async () => {
    const invalid = token({ aud: "https://graph.microsoft.com" });
    const fallback = token();
    let loads = 0;
    let playwrightCalls = 0;
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => state(++loads === 1 ? invalid : fallback),
      acquireSubstrateToken: async () => null,
      fetchTokenWithPlaywright: async () => { playwrightCalls += 1; },
    }});
    expect(await provider.resolveAuthorizationHeader(null)).toBe(`Bearer ${fallback}`);
    expect(playwrightCalls).toBe(1);
  });

  test("persists a valid MSAL result and skips Playwright", async () => {
    const value = token();
    let saved = "";
    let playwrightCalls = 0;
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => null,
      acquireSubstrateToken: async (options) => {
        expect(options.tokenPath).toBe("/token.json");
        return { token: value, expiresAtUtc: new Date(Date.now() + 3_600_000), oid: "object-a", tid: "tenant-a" };
      },
      saveToken: async (_path, candidate) => { saved = candidate; },
      fetchTokenWithPlaywright: async () => { playwrightCalls += 1; },
    }});
    expect(await provider.resolveAuthorizationHeader(null)).toBe(`Bearer ${value}`);
    expect(saved).toBe(value);
    expect(playwrightCalls).toBe(0);
  });

  test("falls back when MSAL persistence fails", async () => {
    const msal = token();
    const fallback = token({ oid: "object-b" });
    let loads = 0;
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => ++loads === 1 ? null : state(fallback),
      acquireSubstrateToken: async () => ({ token: msal, expiresAtUtc: new Date(Date.now() + 3_600_000), oid: "object-a", tid: "tenant-a" }),
      saveToken: async () => { throw new Error("persistence failed"); },
      fetchTokenWithPlaywright: async () => {},
    }});
    expect(await provider.resolveAuthorizationHeader(null)).toBe(`Bearer ${fallback}`);
  });

  test("shares one in-flight acquisition across concurrent callers", async () => {
    const value = token();
    let acquisitions = 0;
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const provider = new ProxyTokenProvider({ dependencies: {
      getTokenPath: async () => "/token.json",
      getBrowserStatePath: async () => "/browser.json",
      loadToken: async () => null,
      acquireSubstrateToken: async () => {
        acquisitions += 1;
        await gate;
        return { token: value, expiresAtUtc: new Date(Date.now() + 3_600_000), oid: "object-a", tid: "tenant-a" };
      },
      saveToken: async () => {},
    }});
    const first = provider.resolveAuthorizationHeader(null);
    const second = provider.resolveAuthorizationHeader(null);
    await Promise.resolve();
    release();
    expect(await Promise.all([first, second])).toEqual([`Bearer ${value}`, `Bearer ${value}`]);
    expect(acquisitions).toBe(1);
  });
});
