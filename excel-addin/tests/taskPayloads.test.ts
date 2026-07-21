import { describe, expect, it } from "vitest";
import {
  buildChartTaskPayload,
  buildCleanTaskPayload,
  buildFormulaTaskPayload,
  buildOcrTaskPayload,
  buildReportTaskPayload,
  isAcceptedOcrFile,
  isImageOcrFile,
  isPdfOcrFile,
} from "../shared/tasks";
import { resolveExcelPromptScenarios } from "../shared/prompts/promptRouting";
import { composeExcelSystemPrompt } from "../shared/prompts/composeExcelPrompt";
import { encodeOpenAiChatContent } from "../shared/provider/messageContentEncode";
import type { AgentMessage } from "../shared/agent/types";

describe("task payloads (desktop-aligned modules)", () => {
  it("formula payload hits formula scenario", () => {
    const payload = buildFormulaTaskPayload({
      dataSourceRanges: ["Sheet1!A1:A10"],
      dataSourceInput: "",
      referenceSampleRange: "B1",
      referenceSampleMode: "partial",
      outputRange: "C1",
      hostEnvironment: "wps",
      task: "求和",
    });
    expect(payload).toContain("【功能模块：生成公式】");
    expect(payload).toContain("数据源选区：Sheet1!A1:A10");
    const scenarios = resolveExcelPromptScenarios({ content: payload });
    expect(scenarios.has("formula")).toBe(true);
  });

  it("clean/chart/report modules route correctly", () => {
    const clean = buildCleanTaskPayload({
      range: "A1:D20",
      task: "去空去重",
      modes: ["drop_empty", "dedupe"],
    });
    expect(clean).toContain("【功能模块：数据清洗】");
    expect(clean).toMatch(/Power Query/);
    const cleanSc = resolveExcelPromptScenarios({ content: clean });
    expect(cleanSc.has("general-office") || cleanSc.has("office-tools")).toBe(
      true,
    );

    const chart = buildChartTaskPayload({
      range: "A1:B5",
      task: "柱状图",
      chartType: "column",
      title: "销售",
      showLegend: true,
      positionNote: "",
    });
    expect(chart).toContain("【功能模块：图表制作】");
    expect(chart).toMatch(/unsupported/);
    const chartSc = resolveExcelPromptScenarios({ content: chart });
    expect(chartSc.has("general-office") || chartSc.has("office-tools")).toBe(
      true,
    );

    const report = buildReportTaskPayload({
      range: "A1:C10",
      task: "周报",
      outputFormat: "excel",
    });
    expect(report).toContain("【功能模块：报告生成】");
    expect(report).toContain("报告工作表");
    expect(report).not.toContain("存储路径");
    const reportWord = buildReportTaskPayload({
      range: "A1",
      task: "x",
      outputFormat: "word",
    });
    expect(reportWord).toMatch(/不适用于加载项|unsupported/);
  });

  it("ocr payload routes ocr-invoice and adapted prompt has no parseDocument how-to as capability", () => {
    const payload = buildOcrTaskPayload({
      mode: "invoice",
      fileNames: ["inv.png"],
      task: "抽字段",
      outputRange: "A1",
    });
    expect(payload).toContain("【功能模块：OCR识别】");
    expect(payload).toContain("WENGGE_OCR_RESULT_V1");
    const sc = resolveExcelPromptScenarios({ content: payload });
    expect(sc.has("ocr-invoice")).toBe(true);
    const prompt = composeExcelSystemPrompt({
      routing: { content: payload },
      officeConnectionStatus: "connected (wps-jsa)",
    });
    expect(prompt).toMatch(/OCR|发票/);
    expect(prompt).toMatch(/不要.*ocr\.parseDocument|无 `ocr\.parseDocument`|无 ocr\.parseDocument/i);
    // hard boundary still forbids desktop stacks
    expect(prompt).toMatch(/COM \/ \.NET Worker \/ Electron IPC/);
  });
});

describe("ocr file acceptance", () => {
  it("accepts images and pdfs; classifies mime", () => {
    const png = new File([new Uint8Array([1])], "a.png", { type: "image/png" });
    const pdf = new File([new Uint8Array([1])], "b.pdf", {
      type: "application/pdf",
    });
    expect(isAcceptedOcrFile(png)).toBe(true);
    expect(isImageOcrFile(png)).toBe(true);
    expect(isPdfOcrFile(pdf)).toBe(true);
    expect(isImageOcrFile(pdf)).toBe(false);
  });
});

describe("multimodal encode (no secret leak in text content)", () => {
  it("encodes image parts for OpenAI chat without putting base64 in role text alone", () => {
    const msg: AgentMessage = {
      role: "user",
      content: "识别这张图",
      contentParts: [
        { type: "image", mimeType: "image/png", base64: "QUJD", fileName: "x.png" },
      ],
    };
    const encoded = encodeOpenAiChatContent(msg);
    expect(Array.isArray(encoded)).toBe(true);
    const json = JSON.stringify(encoded);
    expect(json).toContain("image_url");
    expect(json).toContain("data:image/png;base64,QUJD");
    // API keys never appear
    expect(json).not.toMatch(/sk-[a-zA-Z0-9]/);
  });
});
