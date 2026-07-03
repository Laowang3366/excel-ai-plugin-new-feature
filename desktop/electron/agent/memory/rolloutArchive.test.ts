import { gunzip } from "zlib";
import * as zlib from "zlib";
import { promisify } from "util";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it } from "vitest";

import {
  archiveRolloutSnapshotIfNeeded,
  searchCompressedRolloutMatches,
  spawnRolloutCompressionWorker,
} from "./rolloutArchive";

const gunzipAsync = promisify(gunzip);

describe("archiveRolloutSnapshotIfNeeded", () => {
  it("creates a gzip snapshot when rollout exceeds the configured byte size", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "rollout-archive-"));
    try {
      const filePath = path.join(tempDir, "rollout-thread-1.jsonl");
      await writeFile(filePath, "line-1\nline-2\n", "utf-8");

      const result = await archiveRolloutSnapshotIfNeeded(filePath, {
        archiveRolloutAfterBytes: 4,
      });

      expect(result.archived).toBe(true);
      if (!result.archived) throw new Error("expected archive result");
      expect(result.archivePath).toBe(`${filePath}.gz`);
      const compressed = await readFile(result.archivePath);
      const restored = await gunzipAsync(compressed);
      expect(restored.toString("utf-8")).toBe("line-1\nline-2\n");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("skips small rollout files", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "rollout-archive-"));
    try {
      const filePath = path.join(tempDir, "rollout-thread-1.jsonl");
      await writeFile(filePath, "tiny", "utf-8");

      const result = await archiveRolloutSnapshotIfNeeded(filePath, {
        archiveRolloutAfterBytes: 10_000,
      });

      expect(result).toEqual({ archived: false, reason: "below_threshold" });
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});

describe("spawnRolloutCompressionWorker", () => {
  it("zstd-compresses cold rollout JSONL files and skips active threads", async () => {
      const tempDir = await mkdtemp(path.join(os.tmpdir(), "rollout-compress-"));
    try {
      const coldPath = path.join(tempDir, "2026", "06", "28", "rollout-old-thread-cold.jsonl");
      const activePath = path.join(tempDir, "2026", "06", "29", "rollout-new-thread-active.jsonl");
      await mkdir(path.dirname(coldPath), { recursive: true });
      await mkdir(path.dirname(activePath), { recursive: true });
      await writeFile(coldPath, "cold-line-1\ncold-line-2\n", "utf-8");
      await writeFile(activePath, "active-line\n", "utf-8");

      const oldTime = new Date("2026-06-28T00:00:00Z");
      await import("fs/promises").then((fs) => fs.utimes(coldPath, oldTime, oldTime));

      const result = await spawnRolloutCompressionWorker({
        sessionsRoot: tempDir,
        activeThreadIds: ["thread-active"],
        coldAfterMs: 60_000,
        minBytes: 1,
        now: () => new Date("2026-06-29T00:00:00Z").getTime(),
      });

      expect(result.compressed).toEqual([
        expect.objectContaining({
          sourcePath: coldPath,
          archivePath: `${coldPath}.zst`,
        }),
      ]);
      expect(result.skipped).toEqual([
        expect.objectContaining({
          sourcePath: activePath,
          reason: "active_thread",
        }),
      ]);

      const restored = await zstdDecompressAsync(await readFile(`${coldPath}.zst`));
      expect(restored.toString("utf-8")).toBe("cold-line-1\ncold-line-2\n");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});

describe("searchCompressedRolloutMatches", () => {
  it("searches zstd-compressed rollout JSONL files", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "rollout-search-compressed-"));
    try {
      const sourcePath = path.join(tempDir, "2026", "06", "28", "rollout-old-thread-compressed.jsonl");
      await mkdir(path.dirname(sourcePath), { recursive: true });
      const line = {
        timestamp: "2026-06-28T00:00:00.000Z",
        item: {
          type: "turn_item",
          turnId: "turn-compressed",
          item: {
            type: "user_message",
            id: "msg-compressed",
            content: "compressed archive wellness lesson",
            timestamp: 100,
          },
        },
      };
      const archivePath = `${sourcePath}.zst`;
      await writeFile(archivePath, await zstdCompressAsync(Buffer.from(`${JSON.stringify(line)}\n`, "utf-8")));

      const matches = await searchCompressedRolloutMatches({
        sessionsRoot: tempDir,
        query: "wellness lesson",
        limit: 10,
      });

      expect(matches).toEqual([
        expect.objectContaining({
          threadId: "thread-compressed",
          turnId: "turn-compressed",
          itemType: "turn_item",
          archivePath,
        }),
      ]);
      expect(matches[0].snippet).toContain("wellness lesson");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});

async function zstdCompressAsync(content: Buffer): Promise<Buffer> {
  const compress = (zlib as typeof zlib & {
    zstdCompress?: (buffer: Buffer, callback: (error: Error | null, result: Buffer) => void) => void;
  }).zstdCompress;
  if (!compress) {
    throw new Error("当前 Node 运行时不支持 zstd 压缩");
  }

  return new Promise((resolve, reject) => {
    compress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}

async function zstdDecompressAsync(content: Buffer): Promise<Buffer> {
  const decompress = (zlib as typeof zlib & {
    zstdDecompress?: (buffer: Buffer, callback: (error: Error | null, result: Buffer) => void) => void;
  }).zstdDecompress;
  if (!decompress) {
    throw new Error("当前 Node 运行时不支持 zstd 解压");
  }

  return new Promise((resolve, reject) => {
    decompress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}
