import { afterEach, describe, expect, it, vi } from "vitest";

import {
  resetKnowledgeRegistry,
  setKnowledgeRetriever,
  setKnowledgeStore,
  setKnowledgeWriter,
} from "../../knowledge/knowledgeRegistry";
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

  it("lists indexed knowledge sources from the current registered store", async () => {
    const sources = [
      {
        sourcePath: "D:\\knowledge\\formula-rules.md",
        sourceName: "formula-rules.md",
        sourceType: "md",
        entryCount: 2,
        firstIndexed: 100,
        lastIndexed: 200,
        fileHash: "abc",
      },
    ];
    const executors = new Map();
    addKnowledgeExecutors(executors, {});
    setKnowledgeStore({ listSources: vi.fn(() => sources) } as any);

    const result = await executors.get("knowledge.listSources").execute({});

    expect(result).toEqual({
      success: true,
      data: {
        message: "已读取知识库来源列表",
        sources,
      },
    });
  });

  it("updates an existing knowledge source through the current registered writer", async () => {
    const currentWriter = {
      updateSource: vi.fn(async () => ({
        sourcePath: "D:\\knowledge\\formula-rules.md",
        sourceName: "formula-rules.md",
        entryCount: 1,
        entryIds: ["entry-1"],
        indexedAt: 456,
      })),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, {});
    setKnowledgeWriter(currentWriter as any);

    const result = await executors.get("knowledge.updateSource").execute({
      sourcePath: "D:\\knowledge\\formula-rules.md",
      operation: "append",
      content: "New rule",
      title: "Formula rules",
      tags: ["formula"],
    });

    expect(result).toEqual({
      success: true,
      data: {
        message: "已更新知识库来源",
        sourcePath: "D:\\knowledge\\formula-rules.md",
        sourceName: "formula-rules.md",
        entryCount: 1,
        entryIds: ["entry-1"],
        indexedAt: 456,
      },
    });
    expect(currentWriter.updateSource).toHaveBeenCalledWith({
      sourcePath: "D:\\knowledge\\formula-rules.md",
      operation: "append",
      content: "New rule",
      title: "Formula rules",
      tags: ["formula"],
      metadata: { source: "tool:knowledge.updateSource" },
    });
  });

  it("deletes knowledge source index content through the current registered writer", async () => {
    const currentWriter = {
      deleteSource: vi.fn(async () => ({ sourcePath: "D:\\knowledge\\old.md" })),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, {});
    setKnowledgeWriter(currentWriter as any);

    const result = await executors.get("knowledge.deleteSource").execute({
      sourcePath: "D:\\knowledge\\old.md",
    });

    expect(result).toEqual({
      success: true,
      data: {
        message: "已删除知识库来源索引内容",
        sourcePath: "D:\\knowledge\\old.md",
      },
    });
    expect(currentWriter.deleteSource).toHaveBeenCalledWith({
      sourcePath: "D:\\knowledge\\old.md",
    });
  });
});
