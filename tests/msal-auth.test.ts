import { describe, expect, test } from "bun:test";
import { chmodSync, mkdtempSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { acquireSubstrateToken } from "../src/cli/msal-auth";
import { saveToken } from "../src/cli/token-helpers";

const future = () => Math.floor(Date.now() / 1000) + 3_600;
const jwt = (overrides: Record<string, unknown> = {}) => {
  const encode = (value: unknown) => Buffer.from(JSON.stringify(value)).toString("base64url");
  return `${encode({ alg: "none" })}.${encode({
    aud: "https://substrate.office.com",
    exp: future(),
    tid: "tenant-a",
    oid: "object-a",
    scp: "M365Chat.Read sydney.readwrite",
    ...overrides,
  })}.signature`;
};

const account = (tenantId: string, localAccountId: string) => ({
  tenantId,
  localAccountId,
  homeAccountId: `${localAccountId}.${tenantId}`,
  environment: "login.microsoftonline.com",
  username: "redacted@example.invalid",
});

function fakeApp(accounts: ReturnType<typeof account>[], token: string) {
  const selected: string[] = [];
  let deserialized = "";
  return {
    selected,
    deserialized: () => deserialized,
    app: {
      getTokenCache: () => ({
        deserialize: (value: string) => { deserialized = value; },
        serialize: () => "safe-cache-state",
        getAllAccounts: async () => accounts,
      }),
      acquireTokenSilent: async ({ account: chosen }: { account: { localAccountId: string } }) => {
        selected.push(chosen.localAccountId);
        return { accessToken: token, expiresOn: new Date((future() + 60) * 1000) };
      },
      acquireTokenByCode: async () => { throw new Error("not expected"); },
      getAuthCodeUrl: async () => "https://example.invalid",
    },
  };
}

describe("MSAL token acquisition", () => {
  test("silently renews the matching account and persists an owner-only cache", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
    try {
      const tokenPath = path.join(directory, "token.json");
      const cachePath = path.join(directory, "msal-cache.json");
      await saveToken(tokenPath, jwt(), new Date((future() + 60) * 1000), {
        tid: "tenant-a", oid: "object-a",
      });
      writeFileSync(cachePath, "old-cache");
      chmodSync(cachePath, 0o644);
      const fake = fakeApp(
        [account("tenant-b", "object-b"), account("tenant-a", "object-a")],
        jwt(),
      );
      const result = await acquireSubstrateToken({
        tokenPath, cachePath, allowInteractive: false, quiet: true,
        appFactory: () => fake.app as never,
      });
      expect(result?.tid).toBe("tenant-a");
      expect(fake.selected).toEqual(["object-a"]);
      expect(fake.deserialized()).toBe("old-cache");
      expect(statSync(cachePath).mode & 0o777).toBe(0o600);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("fails safely when multiple accounts are ambiguous", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
    try {
      const fake = fakeApp(
        [account("tenant-a", "object-a"), account("tenant-b", "object-b")],
        jwt(),
      );
      const result = await acquireSubstrateToken({
        tokenPath: path.join(directory, "missing-token.json"),
        cachePath: path.join(directory, "missing-cache.json"),
        allowInteractive: false, quiet: true,
        appFactory: () => fake.app as never,
      });
      expect(result).toBeNull();
      expect(fake.selected).toEqual([]);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  for (const [name, claims] of [
    ["wrong audience", { aud: "https://graph.microsoft.com" }],
    ["expired token", { exp: Math.floor(Date.now() / 1000) - 60 }],
    ["missing scope", { scp: "M365Chat.Read" }],
    ["wrong tenant", { tid: "tenant-b" }],
  ] as const) {
    test(`rejects ${name}`, async () => {
      const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
      try {
        const tokenPath = path.join(directory, "token.json");
        await saveToken(tokenPath, jwt(), new Date((future() + 60) * 1000), {
          tid: "tenant-a", oid: "object-a",
        });
        const fake = fakeApp([account("tenant-a", "object-a")], jwt(claims));
        expect(await acquireSubstrateToken({
          tokenPath,
          cachePath: path.join(directory, "cache.json"),
          allowInteractive: false, quiet: true,
          appFactory: () => fake.app as never,
        })).toBeNull();
      } finally {
        rmSync(directory, { recursive: true, force: true });
      }
    });
  }

  test("accepts the observed Sydney resource audience", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
    try {
      const fake = fakeApp(
        [account("tenant-a", "object-a")],
        jwt({ aud: "https://substrate.office.com/sydney" }),
      );
      const result = await acquireSubstrateToken({
        tokenPath: path.join(directory, "missing-token.json"),
        cachePath: path.join(directory, "cache.json"),
        allowInteractive: false,
        quiet: true,
        appFactory: () => fake.app as never,
      });
      expect(result).not.toBeNull();
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("recovers from a corrupt cache without logging its contents", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
    const cachePath = path.join(directory, "cache.json");
    const sensitive = "sensitive-cache-marker";
    try {
      writeFileSync(cachePath, sensitive);
      const fake = fakeApp([account("tenant-a", "object-a")], jwt());
      (fake.app.getTokenCache().deserialize as (value: string) => void) = () => {
        throw new Error("corrupt cache");
      };
      const errors: string[] = [];
      const original = console.error;
      console.error = (...values: unknown[]) => errors.push(values.join(" "));
      try {
        const result = await acquireSubstrateToken({
          tokenPath: path.join(directory, "missing-token.json"), cachePath,
          allowInteractive: false, quiet: true,
          appFactory: () => fake.app as never,
        });
        expect(result).not.toBeNull();
      } finally {
        console.error = original;
      }
      expect(errors.join(" ")).not.toContain(sensitive);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("returns null when cache persistence fails", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-msal-"));
    try {
      const fake = fakeApp([account("tenant-a", "object-a")], jwt());
      expect(await acquireSubstrateToken({
        tokenPath: path.join(directory, "missing-token.json"),
        cachePath: directory,
        allowInteractive: false, quiet: true,
        appFactory: () => fake.app as never,
      })).toBeNull();
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});
