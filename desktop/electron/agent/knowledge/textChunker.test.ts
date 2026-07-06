import { describe, expect, it } from "vitest";
import { TextChunker } from "./textChunker";
import type { RawChunk } from "./documentParser";

function rawChunk(overrides: Partial<RawChunk> = {}): RawChunk {
  return {
    content: "short content",
    sourcePath: "/knowledge/sample.txt",
    sourceName: "sample.txt",
    sourceType: "txt",
    metadata: { label: "fixture" },
    ...overrides,
  };
}

describe("TextChunker", () => {
  it("keeps small content as one chunk with source metadata intact", () => {
    const chunker = new TextChunker(50, 4);

    const [chunk] = chunker.chunk([
      rawChunk({ content: "plain note", metadata: { section: "intro" } }),
    ]);

    expect(chunk).toMatchObject({
      content: "plain note",
      index: 0,
      sourcePath: "/knowledge/sample.txt",
      sourceName: "sample.txt",
      sourceType: "txt",
      metadata: { section: "intro" },
      tokenCount: 3,
    });
  });

  it("splits tabular data every 100 rows and repeats the header", () => {
    const rows = Array.from({ length: 250 }, (_, index) => `row-${index + 1},value-${index + 1}`);
    const chunker = new TextChunker(10, 1);

    const chunks = chunker.chunk([
      rawChunk({
        content: ["name,value", ...rows].join("\n"),
        sourceType: "csv",
      }),
    ]);

    expect(chunks).toHaveLength(3);
    expect(chunks.map((chunk) => chunk.index)).toEqual([0, 1, 2]);
    expect(chunks.every((chunk) => chunk.content.startsWith("name,value\n"))).toBe(true);
    expect(chunks[0].content).toContain("row-100,value-100");
    expect(chunks[1].content).toContain("row-101,value-101");
    expect(chunks[2].content).toContain("row-250,value-250");
  });

  it("splits markdown at level 2-4 headings without dropping section titles", () => {
    const content = [
      "## Formula rules",
      "Use maintainable dynamic ranges.",
      "### Spill checks",
      "Verify the anchor cell and spill range.",
      "#### WPS notes",
      "Avoid hard-coded demo coordinates.",
    ].join("\n\n");
    const chunker = new TextChunker(12, 1);

    const chunks = chunker.chunk([
      rawChunk({
        content,
        sourceType: "md",
        sourceName: "formula.md",
      }),
    ]);

    expect(chunks.map((chunk) => chunk.content.split("\n")[0])).toEqual([
      "## Formula rules",
      "### Spill checks",
      "#### WPS notes",
    ]);
  });

  it("truncates an oversized plain-text paragraph to the configured token budget", () => {
    const chunker = new TextChunker(20, 2);
    const chunks = chunker.chunk([
      rawChunk({
        content: "x".repeat(120),
        sourceType: "txt",
      }),
    ]);

    expect(chunks).toHaveLength(1);
    expect(chunks[0].content).toHaveLength(40);
    expect(chunks[0].tokenCount).toBe(20);
  });

  it("merges adjacent plain-text paragraphs while they fit within the budget", () => {
    const chunker = new TextChunker(30, 2);
    const chunks = chunker.chunk([
      rawChunk({
        content: ["alpha beta", "gamma delta", "epsilon zeta"].join("\n\n"),
      }),
    ]);

    expect(chunks).toHaveLength(1);
    expect(chunks[0].content).toBe("alpha beta\n\ngamma delta\n\nepsilon zeta");
  });
});
