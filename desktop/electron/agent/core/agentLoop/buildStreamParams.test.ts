import { afterEach, describe, expect, it, vi } from "vitest";

import {
  appendRuntimeDateContext,
  buildEffectiveSystemPrompt,
  appendLongTermMemoryContext,
  appendRuntimeLongTermMemoryContext,
} from "./buildStreamParams";
import * as buildStreamParams from "./buildStreamParams";
import { resetKnowledgeRegistry, setKnowledgeRetriever } from "../../knowledge/knowledgeRegistry";

afterEach(() => {
  resetKnowledgeRegistry();
});

describe("buildStreamParams exports", () => {
  it("does not expose a no-op reasoning mode adapter", () => {
    expect("getEffectiveReasoningMode" in buildStreamParams).toBe(false);
  });
});

describe("appendRuntimeDateContext", () => {
  it("injects the current Shanghai date for relative-time tasks", () => {
    const prompt = appendRuntimeDateContext("base", new Date("2026-07-03T04:30:00.000Z"));

    expect(prompt).toContain("## 运行时上下文");
    expect(prompt).toContain("当前日期：2026");
    expect(prompt).toContain("07");
    expect(prompt).toContain("03");
    expect(prompt).toContain("Asia/Shanghai");
    expect(prompt).toContain("近 N 日");
    expect(prompt).toContain("不要自行补入过期年份");
  });
});

describe("buildEffectiveSystemPrompt", () => {
  it("keeps ordinary Q&A effective prompt under budget without long scenarios", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "VLOOKUP 怎么用？",
    });

    expect(prompt.length).toBeLessThan(6_000);
    expect(prompt).toContain("Office 连接预检铁律");
    expect(prompt).toContain("动态数组函数环境支持：已开启");
    expect(prompt).toContain("不要反复质疑当前环境是否适配动态数组函数");
    expect(prompt).toContain("## 运行时上下文");
    expect(prompt).not.toContain('expand:"spill"');
    expect(prompt).not.toContain('mode:"invoice"');
    expect(prompt).not.toContain("Open XML 优先");
    expect(prompt).not.toContain("### 数据清洗");
  });

  it("injects formula rules for formula assistant turns", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "【功能模块：公式助手】请写入动态数组公式",
    });

    expect(prompt).toContain("场景化操作指南：公式助手");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("禁止为了匹配样例结果硬编码输出路径");
    expect(prompt).toContain("只需更改数据源选区/表引用即可重算");
    expect(prompt).toContain('expand:"spill"');
    expect(prompt.length).toBeLessThan(10_000);
  });

  it("injects OCR rules for invoice turns", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "【功能模块：发票识别】识别字段并写入 Excel",
    });

    expect(prompt).toContain("场景化操作指南：OCR 与发票识别");
    expect(prompt).toContain("ocr.parseDocument");
    expect(prompt).toContain('mode:"invoice"');
    expect(prompt).toContain("发票号码");
  });

  it("injects Office/Open XML rules for Office attachments", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "帮我美化这个文件",
      attachments: [
        { fileName: "demo.pptx", filePath: "D:\\work\\demo.pptx", fileType: "document" },
      ],
    });

    expect(prompt).toContain("Office 工具调用硬性边界");
    expect(prompt).toContain("Open XML 优先");
    expect(prompt).toContain("office.action.apply");
  });

  it("does not pre-inject knowledge context for formula tasks before data is read", async () => {
    const search = vi.fn(async () => [
      {
        entry: {
          sourceName: "formula-rules.md",
          sourcePath: "D:\\kb\\formula-rules.md",
          metadata: {},
          content: "区域汇总公式应该按知识库场景使用 SUMIFS，并以区域字段作为条件。",
        },
        score: 0.92,
      },
    ]);
    setKnowledgeRetriever({
      search,
      formatForPrompt: (results: any[]) => `## 相关知识\n- ${results[0].entry.content}`,
    } as any);

    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "请根据区域汇总公式场景生成 Excel 公式",
    });

    expect(search).not.toHaveBeenCalled();
    expect(prompt).not.toContain("SUMIFS");
    expect(prompt).toContain("场景化操作指南：公式助手");
    expect(prompt).toContain("先用 `office.connection.status`");
    expect(prompt).toContain("读取公式助手提供的数据源选区");
    expect(prompt).toContain("用场景摘要调用 `knowledge.search`");
  });

  it("does not pre-inject knowledge context for Word writing tasks before scene difficulty is known", async () => {
    const search = vi.fn(async () => [
      {
        entry: {
          sourceName: "project-summary.docx",
          sourcePath: "D:\\kb\\project-summary.docx",
          metadata: {},
          content: "项目总结必须包含背景、过程、结果和风险四个章节。",
        },
        score: 0.9,
      },
    ]);
    setKnowledgeRetriever({
      search,
      formatForPrompt: (results: any[]) => `## 相关知识\n- ${results[0].entry.content}`,
    } as any);

    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "根据知识库资料写 Word 项目总结",
      attachments: [
        { fileName: "项目资料.docx", filePath: "D:\\work\\项目资料.docx", fileType: "document" },
      ],
    });

    expect(search).not.toHaveBeenCalled();
    expect(prompt).not.toContain("## 相关知识");
    expect(prompt).not.toContain("背景、过程、结果和风险");
    expect(prompt).toContain("Word 文档、报告、方案");
  });
});

describe("appendLongTermMemoryContext", () => {
  it("injects only user-visible memory", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-user",
        namespace: "global",
        kind: "preference",
        visibility: "user",
        status: "active",
        content: "回复先给结论",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-internal",
        namespace: "global",
        kind: "tool_success_profile",
        visibility: "internal",
        status: "active",
        content: "内部工具统计",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("回复先给结论");
    expect(prompt).not.toContain("内部工具统计");
    expect(prompt).not.toContain("tool_success_profile");
  });

  it("injects only the six office memory kinds", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-project",
        namespace: "global",
        kind: "project_fact",
        visibility: "user",
        status: "active",
        content: "项目事实不应进普通提示词",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-workflow",
        namespace: "global",
        kind: "workflow",
        visibility: "user",
        status: "active",
        content: "工作流不应进普通提示词",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-file",
        namespace: "global",
        kind: "file_impression",
        visibility: "user",
        status: "active",
        content: "这个文件常用数据透视表",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("这个文件常用数据透视表");
    expect(prompt).not.toContain("项目事实不应进普通提示词");
    expect(prompt).not.toContain("工作流不应进普通提示词");
    expect(prompt).not.toContain("project_fact");
    expect(prompt).not.toContain("workflow");
  });

  it("returns the original prompt when there are no active user-visible memories", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-internal",
        namespace: "global",
        kind: "tool_success_profile",
        visibility: "internal",
        status: "active",
        content: "内部工具统计",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toBe("base");
  });

  it("does not inject stale user memory", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-stale",
        namespace: "global",
        kind: "preference",
        visibility: "user",
        status: "stale",
        content: "过期偏好",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toBe("base");
  });

  it("prefers summary over content when present", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-summary",
        namespace: "global",
        kind: "style_preference",
        visibility: "user",
        status: "active",
        content: "很长的原始记忆内容",
        summary: "简短摘要",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("- [style_preference] 简短摘要");
    expect(prompt).not.toContain("很长的原始记忆内容");
  });
});

describe("appendRuntimeLongTermMemoryContext", () => {
  it("loads active user memory from runtime and filters internal/non-office kinds", async () => {
    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async (options) => {
        expect(options).toMatchObject({
          visibility: "user",
          status: "active",
          limit: 8,
        });
        return [
          {
            memoryId: "mem-pref",
            namespace: "global",
            kind: "preference",
            visibility: "user",
            status: "active",
            content: "先给结论",
            createdAt: 1,
            updatedAt: 1,
          },
          {
            memoryId: "mem-tool",
            namespace: "global",
            kind: "tool_success_profile",
            visibility: "internal",
            status: "active",
            content: "内部工具画像",
            createdAt: 1,
            updatedAt: 1,
          },
          {
            memoryId: "mem-project",
            namespace: "global",
            kind: "project_fact",
            visibility: "user",
            status: "active",
            content: "项目事实",
            createdAt: 1,
            updatedAt: 1,
          },
        ];
      },
    });

    expect(prompt).toContain("先给结论");
    expect(prompt).not.toContain("内部工具画像");
    expect(prompt).not.toContain("项目事实");
  });

  it("continues past legacy non-office pages to collect allowed memories", async () => {
    const calls: Array<{ limit?: number; offset?: number }> = [];
    const legacyPage = Array.from({ length: 8 }, (_, index) => ({
      memoryId: `legacy-${index}`,
      namespace: "global",
      kind: index % 2 === 0 ? "project_fact" as const : "workflow" as const,
      visibility: "user" as const,
      status: "active" as const,
      content: `legacy page memory ${index}`,
      createdAt: index,
      updatedAt: index,
    }));

    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async (options) => {
        const listOptions = options ?? {};
        calls.push({ limit: listOptions.limit, offset: listOptions.offset });
        if (listOptions.offset === 0 || listOptions.offset === undefined) {
          return legacyPage;
        }
        return [
          {
            memoryId: "allowed-pref",
            namespace: "global",
            kind: "preference",
            visibility: "user",
            status: "active",
            content: "合法偏好在第二页",
            createdAt: 10,
            updatedAt: 10,
          },
        ];
      },
    });

    expect(calls).toEqual([
      { limit: 8, offset: 0 },
      { limit: 8, offset: 8 },
    ]);
    expect(prompt).toContain("合法偏好在第二页");
    expect(prompt).not.toContain("legacy page memory");
    expect(prompt).not.toContain("project_fact");
    expect(prompt).not.toContain("workflow");
  });

  it("keeps the base prompt when runtime memory loading fails", async () => {
    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async () => {
        throw new Error("db unavailable");
      },
    });

    expect(prompt).toBe("base");
  });
});
