import { describe, expect, test } from "bun:test";
import { mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { loadWrapperOptions } from "../src/proxy/config";
import { SimulatedOutputProtocols } from "../src/proxy/types";

const fixtureRoot = path.join(process.cwd(), "tests");
let fixtureCounter = 0;

describe.serial("simulated output protocol configuration", () => {
  test("defaults to legacy", async () => {
    await withFixture(null, (cwd) =>
      withProtocolEnv(undefined, async () => {
        const options = await loadWrapperOptions(cwd);

        expect(options.simulatedOutputProtocol).toBe(
          SimulatedOutputProtocols.Legacy,
        );
      }),
    );
  });

  test("loads an explicit bridge_v1 value", async () => {
    await withFixture(
      { simulatedOutputProtocol: SimulatedOutputProtocols.BridgeV1 },
      (cwd) =>
        withProtocolEnv(undefined, async () => {
          const options = await loadWrapperOptions(cwd);

          expect(options.simulatedOutputProtocol).toBe(
            SimulatedOutputProtocols.BridgeV1,
          );
        }),
    );
  });

  test("applies CONFIG__simulatedOutputProtocol", async () => {
    await withFixture(
      { simulatedOutputProtocol: SimulatedOutputProtocols.Legacy },
      (cwd) =>
        withProtocolEnv(SimulatedOutputProtocols.BridgeV1, async () => {
          const options = await loadWrapperOptions(cwd);

          expect(options.simulatedOutputProtocol).toBe(
            SimulatedOutputProtocols.BridgeV1,
          );
        }),
    );
  });

  test("rejects an unsupported value", async () => {
    await withFixture(
      { simulatedOutputProtocol: "unsupported" },
      (cwd) =>
        withProtocolEnv(undefined, async () => {
          await expect(loadWrapperOptions(cwd)).rejects.toThrow();
        }),
    );
  });
});

async function withFixture(
  config: Record<string, unknown> | null,
  callback: (cwd: string) => Promise<void>,
): Promise<void> {
  const cwd = path.join(
    fixtureRoot,
    `.config-loading-${process.pid}-${fixtureCounter++}`,
  );
  await mkdir(cwd, { recursive: true });
  if (config) {
    await writeFile(
      path.join(cwd, "config.json"),
      `${JSON.stringify(config)}\n`,
      "utf8",
    );
  }

  try {
    await callback(cwd);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
}

async function withProtocolEnv<T>(
  value: string | undefined,
  callback: () => Promise<T>,
): Promise<T> {
  const key = "CONFIG__simulatedOutputProtocol";
  const previous = process.env[key];
  if (value === undefined) {
    delete process.env[key];
  } else {
    process.env[key] = value;
  }

  try {
    return await callback();
  } finally {
    if (previous === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = previous;
    }
  }
}
