import { afterEach, describe, expect, it, vi } from "vitest";

import { resetKnowledgeRegistry, setKnowledgeRetriever, setKnowledgeWriter } from "../../knowledge/knowledgeRegistry";
import { addKnowledgeExecutors } from "./knowledgeExecutors";

describe("knowledge executors", () => {
  afterEach(() => {
    resetKnowledgeRegistry();
  });

  it("uses the current registered retriever after runtime reload", async () => {
    const staleRetriever = {
      search: vi.fn(async () => []),
      formatForToolResult: vi.fn(() => "stale"),
    };
    const currentRetriever = {
      search: vi.fn(async () => [{ entry: { content: "fresh" }, score: 1 }]),
      formatForToolResult: vi.fn(() => "fresh"),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeRetriever: staleRetriever as any });
    setKnowledgeRetriever(currentRetriever as any);

    const result = await executors.get("knowledge.search").execute({ query: "demo" });

    expect(result).toEqual({ success: true, data: "fresh" });
    expect(currentRetriever.search).toHaveBeenCalledWith({ text: "demo", topK: 5 });
    expect(staleRetriever.search).not.toHaveBeenCalled();
  });

  it("writes notes through the current registered knowledge writer", async () => {
    const staleWriter = { writeNote: vi.fn() };
    const currentWriter = {
      writeNote: vi.fn(async () => ({
        sourcePath: "D:\\knowledge\\note.md",
        sourceName: "note.md",
        entryCount: 1,
        entryIds: ["entry-1"],
        indexedAt: 123,
      })),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeWriter: staleWriter as any });
    setKnowledgeWriter(currentWriter as any);

    const result = await executors.get("knowledge.write").execute({
      title: "字段口径",
      content: "销售额字段必须扣除退款金额。",
      tags: ["销售"],
    });

    expect(result).toEqual({
      success: true,
      data: {
        message: "已写入知识库",
        sourcePath: "D:\\knowledge\\note.md",
        sourceName: "note.md",
        entryCount: 1,
        entryIds: ["entry-1"],
        indexedAt: 123,
      },
    });
    expect(currentWriter.writeNote).toHaveBeenCalledWith({
      title: "字段口径",
      content: "销售额字段必须扣除退款金额。",
      tags: ["销售"],
      sourceName: undefined,
      metadata: { source: "tool:knowledge.write" },
    });
    expect(staleWriter.writeNote).not.toHaveBeenCalled();
  });
});
