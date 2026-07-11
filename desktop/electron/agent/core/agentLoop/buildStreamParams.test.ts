import { afterEach, describe, expect, it, vi } from "vitest";

import {
  buildEffectiveSystemPrompt,
  appendLongTermMemoryContext,
  appendRuntimeLongTermMemoryContext,
} from "./buildStreamParams";
import * as buildStreamParams from "./buildStreamParams";
import { resetKnowledgeRegistry, setKnowledgeRetriever } from "../../knowledge/knowledgeRegistry";
import { buildSystemPrompt } from "../../prompts/systemPrompt";

afterEach(() => {
  resetKnowledgeRegistry();
});

describe("buildStreamParams exports", () => {
  it("does not expose a no-op reasoning mode adapter", () => {
    expect("getEffectiveReasoningMode" in buildStreamParams).toBe(false);
  });
});

describe("buildEffectiveSystemPrompt", () => {
  it("keeps ordinary Q&A effective prompt under budget without long scenarios", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "VLOOKUP 怎么用？",
    });

    expect(prompt.length).toBeLessThan(6_000);
    expect(prompt.startsWith(buildSystemPrompt())).toBe(true);
    expect(prompt).toContain("Office 连接预检铁律");
    expect(prompt).toContain("动态数组函数环境支持：已开启");
    expect(prompt).toContain("版本号或模型记忆中的发布时间都不是函数兼容性证据");
    expect(prompt).toContain("#NAME? 只是名称解析失败的诊断信号");
    expect(prompt).toContain("## 运行时上下文");
    expect(prompt.indexOf("## 运行时上下文")).toBeGreaterThan(
      prompt.indexOf("权限、脚本与质量底线"),
    );
    expect(prompt).not.toContain('expand:"spill"');
    expect(prompt).not.toContain('mode:"invoice"');
    expect(prompt).not.toContain("Open XML 优先");
    expect(prompt).not.toContain("### 数据清洗");
  });

  it("injects formula rules for formula assistant turns", async () => {
    const prompt = await buildEffectiveSystemPrompt(undefined, undefined, {
      content: "【功能模块：公式助手】请写入动态数组公式",
    });

    expect(prompt).toContain("公式助手：内置解题方法");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("参考答案为空时");
    expect(prompt).toContain("选择最小充分公式");
    expect(prompt).toContain("核心方法论已包含在本提示词中");
    expect(prompt).toContain("版本号或函数发布时间推断");
    expect(prompt).toContain("不得声称完成“几步闭环”");
    expect(prompt).toContain("`GROUPBY/PIVOTBY` 能直接完成");
    expect(prompt).toContain("只调用一次或没有消除重复模式的封装才删除");
    expect(prompt).toContain("参数命名服务于读公式");
    expect(prompt).toContain('expand:"spill"');
    expect(prompt).not.toContain("## Office 工具调用硬性边界");
    expect(prompt).not.toContain("场景化操作指南：通用 Office 任务");
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
    expect(prompt).toContain("公式助手：内置解题方法");
    expect(prompt).toContain("office.connection.status -> range.read -> 结构归纳");
    expect(prompt).toContain("必须先看完数据源、参考答案和目标区域");
    expect(prompt).toContain("不要调用知识库来重复加载方法论");
    expect(prompt).toContain("简单任务使用简单公式");
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
  it("loads active user memory from runtime", async () => {
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
        ];
      },
    });

    expect(prompt).toContain("先给结论");
    expect(prompt).not.toContain("内部工具画像");
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
