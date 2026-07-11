import { describe, expect, test } from "vitest";
import {
  appendFolderContext,
  buildContextualPromptSections,
  buildRuntimePromptSection,
  buildSystemPrompt,
} from "./systemPrompt";

describe("buildSystemPrompt", () => {
  test("keeps the lightweight base guardrails", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Office 连接预检铁律");
    expect(prompt).toContain("office.connection.status");
    expect(prompt).not.toMatch(/\{\{[A-Z0-9_]+\}\}/);
    expect(prompt).toContain("range.read");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("ocr.parseDocument");
    expect(prompt).toContain("knowledge.search/write");
    expect(prompt).toContain("memory.search/list/write/delete");
    expect(prompt).toContain("知识库检索时机");
    expect(prompt).toContain("不要在任务开始时只凭用户一句话直接检索知识库");
    expect(prompt).toContain("仅在任务依赖项目资料");
    expect(prompt).not.toContain("公式方法论知识");
    expect(prompt).toContain("shell.execute");
    expect(prompt).toContain("prompt");
    expect(prompt).toContain("forbidden");
  });

  test("keeps concise final reply formatting guidance", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("Markdown 表格");
    expect(prompt).toContain("避免输出原始表格分隔线文本");
    expect(prompt).toContain("不要用 `**`");
  });

  test("does not include long scenario sections by default", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).not.toContain('expand:"spill"');
    expect(prompt).not.toContain("允许输出测试报告");
    expect(prompt).not.toContain("【功能模块：发票识别】");
    expect(prompt).not.toContain('mode:"invoice"');
    expect(prompt).not.toContain("### 数据清洗");
    expect(prompt).not.toContain("### 图表制作");
    expect(prompt).not.toContain("Excel/Word/PPT 高级操作优先使用统一 Office action");
  });

  test("stays within the base prompt budget", () => {
    expect(buildSystemPrompt().length).toBeLessThan(3_000);
  });

  test("keeps shell approval and sandbox guidance", () => {
    const prompt = buildSystemPrompt();

    expect(prompt).toContain("`shell.execute` 受命令安全策略");
    expect(prompt).toContain("prompt");
    expect(prompt).toContain("forbidden");
  });

  test("appends folder context with Office open guidance", () => {
    const prompt = appendFolderContext("base", "D:\\work", "work", [
      { fileName: "demo.xlsx", filePath: "D:\\work\\demo.xlsx", size: 2048 },
    ]);

    expect(prompt).toContain("当前工作文件夹");
    expect(prompt).toContain("demo.xlsx (2KB)");
    expect(prompt).toContain("文件级读取/编辑优先使用 office.action.*");
  });
});

describe("buildRuntimePromptSection", () => {
  test("renders runtime facts after the stable prompt without unresolved variables", () => {
    const prompt = buildRuntimePromptSection({
      officeConnectionStatus: "Word(未连接) | PPT(未连接)",
      dynamicArrayFunctionsEnabled: true,
      now: new Date("2026-07-03T04:30:00.000Z"),
    });

    expect(prompt).toContain("<runtime_context>");
    expect(prompt).toContain("Office 应用连接状态：Word(未连接) | PPT(未连接)");
    expect(prompt).toContain("动态数组函数环境支持：已开启");
    expect(prompt).toContain("当前日期：2026");
    expect(prompt).toContain("Asia/Shanghai");
    expect(prompt).toContain("近 N 日");
    expect(prompt).not.toMatch(/\{\{[A-Z0-9_]+\}\}/);
  });
});

describe("buildContextualPromptSections", () => {
  test("injects formula assistant rules only for formula-writing turns", () => {
    const prompt = buildContextualPromptSections({
      content: "【功能模块：公式助手】请生成动态数组公式并写入答案区域",
    });

    expect(prompt).toContain("公式助手：内置解题方法");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("knowledge.search");
    expect(prompt).toContain("核心方法论已包含在本提示词中");
    expect(prompt).toContain("不要调用知识库来重复加载方法论");
    expect(prompt).toContain("输入形状");
    expect(prompt).toContain("输出形状");
    expect(prompt).toContain("参考答案为空时");
    expect(prompt).toContain("选择最小充分公式");
    expect(prompt).toContain("索引/掩码/状态");
    expect(prompt).toContain("生成业务键、条件或索引");
    expect(prompt).toContain("环境支持且模式稳定");
    expect(prompt).toContain('expand:"spill"');
    expect(prompt).toContain("ASCII");
    expect(prompt).toContain("动态数组只写目标锚点");
    expect(prompt).toContain("写入成功不等于结果正确");
    expect(prompt).not.toContain("## Office 工具调用硬性边界");
    expect(prompt).not.toContain("场景化操作指南：通用 Office 任务");
    expect(prompt.length).toBeLessThan(3_500);
  });

  test("injects OCR and invoice rules for invoice-recognition turns", () => {
    const prompt = buildContextualPromptSections({
      content: "【功能模块：发票识别】识别字段并写入表格",
    });

    expect(prompt).toContain("场景化操作指南：OCR 与发票识别");
    expect(prompt).toContain("ocr.parseDocument");
    expect(prompt).toContain('mode:"invoice"');
    expect(prompt).toContain("发票号码");
    expect(prompt).toContain("range.write");
    expect(prompt).toContain("写入后回读验证一次");
    expect(prompt.length).toBeLessThan(2_000);
  });

  test("injects OCR rules for image or PDF attachments", () => {
    const prompt = buildContextualPromptSections({
      content: "帮我看看这个附件",
      attachments: [
        { fileName: "invoice.pdf", filePath: "D:\\work\\invoice.pdf", fileType: "document" },
      ],
    });

    expect(prompt).toContain("ocr.parseDocument");
    expect(prompt).toContain('mode:"invoice"');
  });

  test("injects Office and Open XML rules for Office file editing", () => {
    const prompt = buildContextualPromptSections({
      content: "美化当前 PPT 版面，并检查视觉效果",
    });

    expect(prompt).toContain("Office 工具调用硬性边界");
    expect(prompt).toContain("Open XML 优先");
    expect(prompt).toContain("office.action.inspect");
    expect(prompt).toContain("office.action.apply");
    expect(prompt).toContain("Word 文档、报告、方案");
    expect(prompt).toContain("判断写作难度");
    expect(prompt).toContain("简单改写或短文本补全不搜库");
    expect(prompt).toContain('office.action.apply({ app, action:"snapshot", operation:"snapshot"');
    expect(prompt).not.toContain("office.action.inspect 获取结构、表格和截图信息");
    expect(prompt).toContain('preferEngine:"com"');
    expect(prompt.length).toBeLessThan(2_500);
  });

  test("injects general scenario rules for data-cleaning and report tasks", () => {
    const prompt = buildContextualPromptSections({
      content: "清洗这张表并生成统计报告",
    });

    expect(prompt).toContain("场景化操作指南：通用 Office 任务");
    expect(prompt).toContain("### 数据清洗");
    expect(prompt).toContain("### 分析报告");
    expect(prompt.length).toBeLessThan(3_500);
  });

  test("does not inject formula, OCR or Open XML long sections for ordinary Q&A", () => {
    const prompt = buildContextualPromptSections({
      content: "VLOOKUP 怎么用？",
    });

    expect(prompt).toBe("");
  });
});
