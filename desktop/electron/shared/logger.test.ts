import { mkdtemp, readFile, rm } from "fs/promises";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("electron", () => ({
  app: {
    getPath: vi.fn(() => os.tmpdir()),
  },
}));

import { configureLogDirectory, createLogger } from "./logger";

const CANARY = "sk-1234567890abcdefghijklmnop";

describe("logger", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("redacts messages and structured fields before writing a single bounded line", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "logger-redaction-"));
    try {
      configureLogDirectory(tempDir);
      vi.spyOn(console, "log").mockImplementation(() => undefined);

      createLogger("security-test").error(`request failed ${CANARY}`, {
        apiKey: CANARY,
        nested: { authorization: `Bearer ${CANARY}` },
        longValue: "x".repeat(30_000),
      });

      const date = new Date().toISOString().split("T")[0];
      const content = await readFile(path.join(tempDir, `app-${date}.log`), "utf-8");

      expect(content).not.toContain(CANARY);
      expect(content).toContain("[REDACTED:openai-style-key]");
      expect(content).toContain('"apiKey":"[REDACTED]"');
      expect(content.trim().length).toBeLessThanOrEqual(16_384);
      expect(content.trim().split(/\r?\n/)).toHaveLength(1);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("does not let hostile structured values break the logging path", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "logger-hostile-"));
    try {
      configureLogDirectory(tempDir);
      vi.spyOn(console, "log").mockImplementation(() => undefined);
      const hostile = Object.defineProperty({}, "value", {
        enumerable: true,
        get: () => {
          throw new Error("getter failed");
        },
      });

      expect(() => createLogger("security-test").error("failed", hostile)).not.toThrow();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
