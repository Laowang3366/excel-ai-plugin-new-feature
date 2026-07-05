import { describe, expect, it } from "vitest";
import {
  buildCodeTaskPayload,
  buildFormulaTaskPayload,
  buildReportTaskPayload,
  normalizeHostEnvironment,
} from "./taskComposerPayloads";

describe("taskComposerPayloads", () => {
  it("builds formula payload with sample mode and current host", () => {
    const payload = buildFormulaTaskPayload({
      dataSourceRanges: ["Sheet1!A1:C10"],
      dataSourceInput: "",
      referenceSampleRange: "Sheet1!E1:G3",
      referenceSampleMode: "complete",
      outputRange: "Sheet1!I1",
      hostEnvironment: "wps",
      task: "按部门汇总金额",
    });

    expect(payload).toContain("【功能模块：生成公式】");
    expect(payload).toContain("当前连接环境：WPS");
    expect(payload).toContain("答案参考样例类型：完整样例");
    expect(payload).toContain("交付要求：必须使用 Excel/WPS 函数公式完成");
    expect(payload).not.toContain("是否支持动态数组");
  });

  it("builds code payload with partial sample and Microsoft Excel host", () => {
    const payload = buildCodeTaskPayload({
      dataSourceRanges: [],
      dataSourceInput: "Sheet1!A1:C10",
      referenceSampleRange: "Sheet1!E1:G3",
      referenceSampleMode: "partial",
      outputRange: "",
      hostEnvironment: "microsoft_excel",
      preferredLanguage: "auto",
      task: "生成录入窗体",
    });

    expect(payload).toContain("【功能模块：代码生成】");
    expect(payload).toContain("运行环境：Microsoft Excel");
    expect(payload).toContain("答案参考样例类型：部分样例");
  });

  it("builds report payloads for document and spreadsheet outputs", () => {
    expect(buildReportTaskPayload({
      range: "Sheet1!A1:F20",
      task: "输出月度经营报告",
      outputFormat: "word",
      storagePath: "C:\\Users\\wfq\\Desktop",
    })).toContain("创建 Word 文档");

    expect(buildReportTaskPayload({
      range: "Sheet1!A1:F20",
      task: "输出汇报 PPT",
      outputFormat: "ppt",
      storagePath: "C:\\Users\\wfq\\Desktop",
    })).toContain("创建 PPT 文件");

    expect(buildReportTaskPayload({
      range: "Sheet1!A1:F20",
      task: "输出表格报告",
      outputFormat: "excel",
      storagePath: "",
    })).toContain("在当前连接的 Excel/WPS 环境中新增或更新报告工作表");
  });

  it("normalizes connected host names", () => {
    expect(normalizeHostEnvironment({ connected: true, host: "wps" })).toBe("wps");
    expect(normalizeHostEnvironment({ connected: true, host: "excel" })).toBe("microsoft_excel");
    expect(normalizeHostEnvironment({ connected: false, host: "" })).toBe("unknown");
  });

});
