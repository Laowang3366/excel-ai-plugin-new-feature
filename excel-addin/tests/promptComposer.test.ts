import { describe, expect, it } from "vitest";
import {
  EXCLUDED_SCENARIOS,
  composeExcelSystemPrompt,
  composePromptSections,
  renderPromptTemplate,
  resolveExcelPromptScenarios,
} from "../shared/prompts";

describe("promptComposer parity", () => {
  it("throws when template variables are missing", () => {
    expect(() => renderPromptTemplate("Hello {{NAME}}", {})).toThrow(
      "缺少提示词模板变量：NAME",
    );
  });

  it("throws on unresolved non-standard placeholders", () => {
    expect(() => renderPromptTemplate("Hi {{name}}", { name: "x" })).toThrow(
      "存在未替换的提示词模板变量",
    );
  });

  it("composes unique non-empty sections", () => {
    const text = composePromptSections([
      { key: "a", content: " one " },
      { key: "a", content: "dup" },
      { key: "b", content: "two" },
      { key: "c", content: "   " },
    ]);
    expect(text).toBe("one\n\ntwo");
  });
});

describe("excel-only prompt routing", () => {
  it("routes formula, office tools, and add-in OCR scenario", () => {
    const formula = resolveExcelPromptScenarios({ content: "请写入动态数组公式" });
    expect(formula.has("formula")).toBe(true);

    const tools = resolveExcelPromptScenarios({ content: "清洗 excel 工作表数据" });
    expect(tools.has("office-tools") || tools.has("general-office")).toBe(true);

    const ocr = resolveExcelPromptScenarios({ content: "【功能模块：OCR识别】识别发票图片" });
    expect(ocr.has("formula")).toBe(false);
    expect(ocr.has("ocr-invoice")).toBe(true);
    expect(EXCLUDED_SCENARIOS).not.toContain("ocr-invoice");
  });

  it("builds excel system prompt from adapted + synced templates", () => {
    const prompt = composeExcelSystemPrompt({
      routing: { content: "【功能模块：公式】写一个 SUM" },
      officeConnectionStatus: "connected (office-js)",
      now: new Date("2026-07-19T04:00:00.000Z"),
      dynamicArrayEnabled: true,
    });
    expect(prompt.length).toBeGreaterThan(100);
    expect(prompt).toContain("Office 应用连接状态：connected (office-js)");
    expect(prompt).toMatch(/2026/);
    expect(prompt).toContain("host.status");
    expect(prompt).toContain("动态数组");
    // formula-only turn should not pull OCR scenario template path marker as capability
    expect(prompt).not.toMatch(/office\.connection\.status|createWorkbook|preferEngine/);
  });
});
