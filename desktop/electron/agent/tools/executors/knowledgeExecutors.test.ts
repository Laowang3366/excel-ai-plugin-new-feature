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

  it("limits formula methodology searches to the builtin methodology source", async () => {
    const retriever = {
      search: vi.fn(async () => []),
      formatForToolResult: vi.fn(() => "methodology"),
    };
    const sources = [
      {
        sourcePath: "D:\\app\\public\\knowledge\\excel-wps-formula-problem-solving-methodology.md",
        sourceName: "excel-wps-formula-problem-solving-methodology.md",
      },
      {
        sourcePath: "D:\\user\\case.xlsx",
        sourceName: "case.xlsx",
      },
    ];
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeRetriever: retriever as any });
    setKnowledgeStore({ listSources: vi.fn(() => sources) } as any);

    const result = await executors.get("knowledge.search").execute({
      query: "按部门分组聚合",
      topK: 6,
      scope: "formula_methodology",
    });

    expect(result).toEqual({ success: true, data: "methodology" });
    expect(retriever.search).toHaveBeenCalledWith({
      text: "按部门分组聚合",
      topK: 6,
      pathFilter: [sources[0].sourcePath],
    });
  });

  it("returns no_match without searching when no scene knowledge source exists", async () => {
    const retriever = {
      search: vi.fn(async () => []),
      formatForToolResult: vi.fn(() => "none"),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeRetriever: retriever as any });
    setKnowledgeStore({
      listSources: vi.fn(() => [{
        sourcePath: "D:\\app\\public\\knowledge\\excel-wps-formula-problem-solving-methodology.md",
        sourceName: "excel-wps-formula-problem-solving-methodology.md",
      }]),
    } as any);

    const result = await executors.get("knowledge.search").execute({
      query: "分组聚合",
      scope: "formula_scene",
    });

    expect(result).toMatchObject({ success: true, data: { status: "no_match", matchCount: 0 } });
    expect(retriever.search).not.toHaveBeenCalled();
  });

  it("searches only user scene sources and reports matched results", async () => {
    const results = [{ entry: { content: "scene" }, score: 0.8 }];
    const retriever = {
      search: vi.fn(async () => results),
      formatForToolResult: vi.fn(() => "scene knowledge"),
    };
    const sources = [
      {
        sourcePath: "D:\\app\\public\\knowledge\\excel-wps-formula-problem-solving-methodology.md",
        sourceName: "excel-wps-formula-problem-solving-methodology.md",
      },
      { sourcePath: "D:\\user\\course.xlsx", sourceName: "course.xlsx" },
    ];
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeRetriever: retriever as any });
    setKnowledgeStore({ listSources: vi.fn(() => sources) } as any);

    const result = await executors.get("knowledge.search").execute({
      query: "分组聚合",
      topK: 3,
      scope: "formula_scene",
    });

    expect(result).toMatchObject({ success: true, data: { status: "matched", matchCount: 1, content: "scene knowledge" } });
    expect(retriever.search).toHaveBeenCalledWith({
      text: "分组聚合",
      topK: 3,
      pathFilter: [sources[1].sourcePath],
    });
  });

  it("reports optional scene retrieval as unavailable instead of failing the workflow", async () => {
    const retriever = {
      search: vi.fn(async () => { throw new Error("embedding offline"); }),
      formatForToolResult: vi.fn(),
    };
    const executors = new Map();
    addKnowledgeExecutors(executors, { knowledgeRetriever: retriever as any });
    setKnowledgeStore({
      listSources: vi.fn(() => [{ sourcePath: "D:\\user\\scene.md", sourceName: "scene.md" }]),
    } as any);

    const result = await executors.get("knowledge.search").execute({
      query: "文本提取",
      scope: "formula_scene",
    });

    expect(result).toMatchObject({ success: true, data: { status: "unavailable", matchCount: 0 } });
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
