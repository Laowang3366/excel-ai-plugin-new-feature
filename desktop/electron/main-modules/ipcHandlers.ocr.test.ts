import { describe, expect, it } from "vitest";

import {
  isLikelyInvoiceFileList,
  normalizeOcrMode,
  normalizePlainOcrText,
} from "./ocrModeDetection";

describe("OCR IPC helpers", () => {
  it("treats electronic invoice filenames as invoice candidates", () => {
    expect(normalizeOcrMode("image")).toBe("image");
    expect(isLikelyInvoiceFileList(["C:/tmp/滴滴电子发票B.pdf"])).toBe(true);
  });

  it("converts plain invoice OCR text into structured invoice fields", () => {
    const result = normalizePlainOcrText("image", [
      "电子发票",
      "发票号码：12345678",
      "开票日期：2026年07月03日",
      "购买方名称：上海示例科技有限公司",
      "销售方名称：北京供应商有限公司",
      "价税合计：106.00",
    ].join("\n"));

    expect(result.kind).toBe("invoice");
    expect(result.fields).toMatchObject({
      发票号码: "12345678",
      开票日期: "2026年07月03日",
      购买方名称: "上海示例科技有限公司",
      销售方名称: "北京供应商有限公司",
      价税合计: "106.00",
    });
    expect(result.rows[0]).toContain("发票号码");
  });
});
