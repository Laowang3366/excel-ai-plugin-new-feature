import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

const { MAX_CHUNK_BYTES, MAX_ENTRY_BYTES, inspectRendererBundle } =
  require("../../scripts/check-renderer-bundle-budget.cjs") as {
    MAX_CHUNK_BYTES: number;
    MAX_ENTRY_BYTES: number;
    inspectRendererBundle: (distDir: string) => {
      entry: { fileName: string; bytes: number };
      asyncChunkCount: number;
    };
  };

const temporaryDirectories: string[] = [];

function createBundle(entryBytes: number, asyncChunkBytes = 128): string {
  const distDir = fs.mkdtempSync(path.join(os.tmpdir(), "renderer-budget-"));
  temporaryDirectories.push(distDir);
  const assetsDir = path.join(distDir, "assets");
  fs.mkdirSync(assetsDir);
  fs.writeFileSync(
    path.join(distDir, "index.html"),
    '<script type="module" crossorigin src="./assets/index-test.js"></script>',
  );
  fs.writeFileSync(path.join(assetsDir, "index-test.js"), Buffer.alloc(entryBytes));
  fs.writeFileSync(path.join(assetsDir, "settings-test.js"), Buffer.alloc(asyncChunkBytes));
  fs.writeFileSync(path.join(assetsDir, "office-test.js"), Buffer.alloc(128));
  return distDir;
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

describe("renderer bundle budget", () => {
  it("accepts a split bundle within the entry and chunk limits", () => {
    const report = inspectRendererBundle(createBundle(MAX_ENTRY_BYTES));
    expect(report.entry.bytes).toBe(MAX_ENTRY_BYTES);
    expect(report.asyncChunkCount).toBe(2);
  });

  it("rejects an oversized initial entry", () => {
    expect(() => inspectRendererBundle(createBundle(MAX_ENTRY_BYTES + 1))).toThrow(
      "Renderer 首屏入口超出预算",
    );
  });

  it("rejects an oversized async chunk", () => {
    expect(() => inspectRendererBundle(createBundle(128, MAX_CHUNK_BYTES + 1))).toThrow(
      "Renderer chunk 超出预算",
    );
  });

  it("does not count eagerly preloaded modules as async chunks", () => {
    const distDir = createBundle(128);
    fs.appendFileSync(
      path.join(distDir, "index.html"),
      '<link rel="modulepreload" href="./assets/settings-test.js">',
    );

    expect(() => inspectRendererBundle(distDir)).toThrow("Renderer 异步 chunk 数量不足");
  });
});
