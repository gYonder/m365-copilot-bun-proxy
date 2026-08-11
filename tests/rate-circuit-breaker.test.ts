import { describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { DurableStateStore } from "../src/proxy/durable-state";
import { RateCircuitBreaker } from "../src/proxy/rate-circuit-breaker";

describe("RateCircuitBreaker", () => {
  test("opens, persists, backs off, and closes after success", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "m365-rate-"));
    try {
      const store = new DurableStateStore(path.join(dir, "state.json"));
      const breaker = new RateCircuitBreaker(store);
      breaker.record(429, 1_000);
      expect(breaker.allow(1_001)).toBe(false);
      expect(breaker.retryAfterSeconds(1_001)).toBe(60);
      expect(JSON.parse(readFileSync(path.join(dir, "state.json"), "utf8")).rateLimit.consecutive).toBe(1);

      breaker.record(429, 2_000);
      expect(breaker.status(2_001).consecutive).toBe(2);
      expect(breaker.status(2_001).openUntilUtc).toBe(122_000);
      breaker.record(200, 123_000);
      expect(breaker.allow(123_001)).toBe(true);
      expect(breaker.status(123_001).consecutive).toBe(0);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test("reads pre-circuit continuation state", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "m365-rate-"));
    try {
      const file = path.join(dir, "state.json");
      writeFileSync(file, JSON.stringify({
        schema_version: 1,
        contract_version: 2,
        adapter_route: "substrate-coding",
        provider_runtime_version: "0.1.0",
        responses: {}, conversations: {}, sessions: {}, replays: {}, toolLedgers: {},
      }));
      const store = new DurableStateStore(file);
      expect(new RateCircuitBreaker(store).status().open).toBe(false);
      expect(store.quarantineStatus.count).toBe(0);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
