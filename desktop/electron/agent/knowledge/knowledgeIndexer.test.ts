import { mkdtemp, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it, vi } from "vitest";

import { KnowledgeIndexer } from "./knowledgeIndexer";
import { SqliteStore } from "./sqliteStore";

function rawChunk(filePath: string, content: string) {
  return {
    content,
    sourcePath: filePath,
    sourceName: path.basename(filePath),
    sourceType: "txt",
    metadata: {},
  };
}

describe("KnowledgeIndexer", () => {
  it("preserves the previous index when embedding generation fails", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "knowledge-indexer-"));
    const filePath = path.join(root, "source.txt");
    await writeFile(filePath, "old content");
    const store = new SqliteStore(":memory:");
    await store.init();
    const parser = { parseAsync: vi.fn(async () => [rawChunk(filePath, "old content")]) };
    const embedder = {
      embedBatch: vi.fn().mockResolvedValueOnce([[1, 0]]),
      getProfile: vi.fn(() => ({ provider: "test", model: "embed", dimensions: 2 })),
    };
    const indexer = new KnowledgeIndexer(store, embedder as any, parser as any);

    try {
      await expect(indexer.indexFile(filePath, { skipUnchanged: false })).resolves.toMatchObject({
        success: true,
      });
      parser.parseAsync.mockResolvedValueOnce([rawChunk(filePath, "new content")]);
      embedder.embedBatch.mockRejectedValueOnce(new Error("embedding unavailable"));

      await expect(indexer.indexFile(filePath, { skipUnchanged: false })).resolves.toMatchObject({
        success: false,
        error: "embedding unavailable",
      });
      expect(store.getEntriesBySource(filePath).map((entry) => entry.content)).toEqual([
        "old content",
      ]);
      expect(store.getEntriesBySource(filePath)[0].embedding).toEqual([1, 0]);
    } finally {
      store.close();
      await rm(root, { recursive: true, force: true });
    }
  });

  it("removes stale entries when a changed file parses to no content", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "knowledge-indexer-"));
    const filePath = path.join(root, "source.txt");
    await writeFile(filePath, "content");
    const store = new SqliteStore(":memory:");
    await store.init();
    const parser = { parseAsync: vi.fn(async () => [rawChunk(filePath, "content")]) };
    const embedder = {
      embedBatch: vi.fn(async () => [[1, 0]]),
      getProfile: vi.fn(() => ({ provider: "test", model: "embed", dimensions: 2 })),
    };
    const indexer = new KnowledgeIndexer(store, embedder as any, parser as any);

    try {
      await indexer.indexFile(filePath, { skipUnchanged: false });
      parser.parseAsync.mockResolvedValueOnce([]);
      await expect(indexer.indexFile(filePath, { skipUnchanged: false })).resolves.toMatchObject({
        success: true,
        entryCount: 0,
      });
      expect(store.getEntriesBySource(filePath)).toEqual([]);
      expect(store.getSource(filePath)).toBeNull();
    } finally {
      store.close();
      await rm(root, { recursive: true, force: true });
    }
  });
});
