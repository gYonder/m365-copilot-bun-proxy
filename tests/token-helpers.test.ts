import { describe, expect, test } from "bun:test";
import {
  chmodSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { loadToken, saveToken } from "../src/cli/token-helpers";

describe("token persistence", () => {
  test("writes token state atomically with owner-only permissions", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-token-"));
    const tokenPath = path.join(directory, "token.json");
    try {
      await saveToken(
        tokenPath,
        "opaque-test-value",
        new Date("2030-01-01T00:00:00.000Z"),
        { oid: "object-test", tid: "tenant-test" },
      );
      expect(statSync(tokenPath).mode & 0o777).toBe(0o600);
      expect(JSON.parse(readFileSync(tokenPath, "utf8"))).toEqual({
        token: "opaque-test-value",
        expiresAtUtc: "2030-01-01T00:00:00.000Z",
        oid: "object-test",
        tid: "tenant-test",
      });
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("repairs overly broad permissions when loading existing state", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-token-"));
    const tokenPath = path.join(directory, "token.json");
    try {
      writeFileSync(
        tokenPath,
        JSON.stringify({
          token: "opaque-test-value",
          expiresAtUtc: "2030-01-01T00:00:00.000Z",
        }),
      );
      chmodSync(tokenPath, 0o644);
      expect(await loadToken(tokenPath)).not.toBeNull();
      expect(statSync(tokenPath).mode & 0o777).toBe(0o600);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("rejects corrupt state while still repairing its permissions", async () => {
    const directory = mkdtempSync(path.join(tmpdir(), "m365-token-"));
    const tokenPath = path.join(directory, "token.json");
    try {
      writeFileSync(tokenPath, "not-json");
      chmodSync(tokenPath, 0o644);
      expect(await loadToken(tokenPath)).toBeNull();
      expect(statSync(tokenPath).mode & 0o777).toBe(0o600);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});
