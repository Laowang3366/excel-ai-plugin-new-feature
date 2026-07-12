import { describe, expect, it } from "vitest";

import { resolvePromptScenarios } from "./promptRouting";

describe("resolvePromptScenarios", () => {
  it("uses the feature module marker as the primary route", () => {
    expect(resolvePromptScenarios({ content: "【功能模块：生成公式】按部门汇总" })).toEqual(
      new Set(["formula"]),
    );
    expect(resolvePromptScenarios({ content: "【功能模块：报告生成】输出经营月报" })).toEqual(
      new Set(["office-tools", "general-office"]),
    );
  });

  it("recognizes direct user intents without loading formula rules for Q&A", () => {
    expect(resolvePromptScenarios({ content: "请生成 Excel 动态数组公式" })).toContain("formula");
    expect(resolvePromptScenarios({ content: "VLOOKUP 怎么用？" })).toEqual(new Set());
  });

  it("routes supported attachments without relying on message keywords", () => {
    expect(
      resolvePromptScenarios({
        content: "处理这个附件",
        attachments: [{ fileName: "invoice.pdf", fileType: "document" }],
      }),
    ).toContain("ocr-invoice");
    expect(
      resolvePromptScenarios({
        content: "处理这个附件",
        attachments: [{ fileName: "report.xlsx", fileType: "document" }],
      }),
    ).toContain("office-tools");
  });

  it("loads the macro workflow only for persistent macro tasks", () => {
    expect(resolvePromptScenarios({ content: "创建 VBA 宏并绑定按钮" })).toContain("macro");
    expect(resolvePromptScenarios({ content: "运行现有宏" })).toContain("macro");
    expect(resolvePromptScenarios({ content: "把几个单元格作为控制按钮，点击按钮切换基础数据" })).toContain("macro");
    expect(resolvePromptScenarios({ content: "用脚本临时整理文件" })).not.toContain("macro");
  });
});
