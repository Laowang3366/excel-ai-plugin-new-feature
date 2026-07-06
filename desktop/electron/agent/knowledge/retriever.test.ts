import { describe, expect, it, vi } from "vitest";
import { Retriever } from "./retriever";
import type { EmbeddingService } from "./embeddingService";
import type { SqliteStore } from "./sqliteStore";
import type { KnowledgeEntry, KnowledgeResult } from "./types";

function entry(overrides: Partial<KnowledgeEntry> = {}): KnowledgeEntry {
  return {
    id: "entry-1",
    source: "document",
    sourcePath: "/knowledge/formula.md",
    sourceName: "formula.md",
    sourceType: "md",
    chunkIndex: 0,
    content: "Formula knowledge",
    metadata: {},
    embedding: [1, 0, 0],
    embeddingProvider: "openai",
    embeddingModel: "text-embedding-3-small",
    embeddingDimensions: 3,
    indexedAt: 1,
    tokenCount: 4,
    ...overrides,
  };
}

function createHarness(options?: { vectorResults?: KnowledgeResult[]; keywordResults?: KnowledgeEntry[] }) {
  const vectorResults = options?.vectorResults ?? [
    { entry: entry({ id: "vector-1", content: "Dynamic array formula guide" }), score: 0.92 },
  ];
  const keywordResults = options?.keywordResults ?? [
    { ...entry({ id: "keyword-1", content: "Keyword fallback formula guide" }) },
  ];

  const store = {
    searchByVector: vi.fn(() => vectorResults),
    searchByKeyword: vi.fn(() => keywordResults),
  };
  const embedder = {
    embed: vi.fn(async () => [1, 0, 0]),
    getProfile: vi.fn(() => ({
      provider: "openai",
      model: "text-embedding-3-small",
      dimensions: 3,
    })),
  };

  const retriever = new Retriever(
    store as unknown as SqliteStore,
    embedder as unknown as EmbeddingService,
    { candidateCount: 20, defaultTopK: 2, minScore: 0.2 }
  );

  return { retriever, store, embedder };
}

describe("Retriever", () => {
  it("passes the active embedding profile and caller filters into vector search", async () => {
    const { retriever, store } = createHarness();

    await retriever.search({
      text: "formula",
      topK: 2,
      sourceFilter: ["document"],
      pathFilter: ["/knowledge/formula.md"],
    });

    expect(store.searchByVector).toHaveBeenCalledWith(
      [1, 0, 0],
      20,
      {
        embeddingProfile: {
          provider: "openai",
          model: "text-embedding-3-small",
          dimensions: 3,
        },
        sourceFilter: ["document"],
        pathFilter: ["/knowledge/formula.md"],
      }
    );
  });

  it("filters low vector scores before applying topK", async () => {
    const { retriever } = createHarness({
      vectorResults: [
        { entry: entry({ id: "high" }), score: 0.95 },
        { entry: entry({ id: "too-low" }), score: 0.19 },
        { entry: entry({ id: "second" }), score: 0.7 },
      ],
    });

    const results = await retriever.search({ text: "formula", topK: 1 });

    expect(results.map((result) => result.entry.id)).toEqual(["high"]);
  });

  it("falls back to keyword search when embedding generation fails", async () => {
    const { retriever, store, embedder } = createHarness();
    embedder.embed.mockRejectedValueOnce(new Error("embedding unavailable"));

    const results = await retriever.search({
      text: "dynamic array formula",
      topK: 2,
      sourceFilter: ["document"],
    });

    expect(store.searchByVector).not.toHaveBeenCalled();
    expect(store.searchByKeyword).toHaveBeenCalledWith(
      expect.arrayContaining(["dynamic array formula", "dynamic", "array", "formula"]),
      2,
      { sourceFilter: ["document"], pathFilter: undefined }
    );
    expect(results).toEqual([{ entry: expect.objectContaining({ id: "keyword-1" }), score: 0 }]);
  });

  it("falls back to keyword search when vector matches are below the minimum score", async () => {
    const { retriever, store } = createHarness({
      vectorResults: [{ entry: entry({ id: "weak-vector" }), score: 0.05 }],
      keywordResults: [entry({ id: "keyword-hit", content: "Formula fallback" })],
    });

    const results = await retriever.search({ text: "formula", topK: 2 });

    expect(store.searchByKeyword).toHaveBeenCalled();
    expect(results).toEqual([{ entry: expect.objectContaining({ id: "keyword-hit" }), score: 0 }]);
  });

  it("formats prompt context grouped by source and keeps empty results silent", () => {
    const { retriever } = createHarness();

    const formatted = retriever.formatForPrompt([
      {
        entry: entry({
          id: "a",
          sourcePath: "/knowledge/sales.xlsx",
          sourceName: "sales.xlsx",
          metadata: { sheetName: "Rules" },
          content: "Use SUMIFS for conditional summaries.",
        }),
        score: 0.8,
      },
      {
        entry: entry({
          id: "b",
          sourcePath: "/knowledge/sales.xlsx",
          sourceName: "sales.xlsx",
          metadata: { sheetName: "Rules" },
          content: "Use FILTER for spill results.",
        }),
        score: 0.7,
      },
    ]);

    expect(retriever.formatForPrompt([])).toBe("");
    expect(formatted).toContain("sales.xlsx");
    expect(formatted.match(/sales\.xlsx/g)).toHaveLength(1);
    expect(formatted).toContain("SUMIFS");
    expect(formatted).toContain("FILTER");
  });

  it("formats tool output with an explicit empty-state message", () => {
    const { retriever } = createHarness();

    expect(retriever.formatForToolResult([])).toContain("未找到");
  });
});
