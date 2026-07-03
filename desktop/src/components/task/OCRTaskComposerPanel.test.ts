import { describe, expect, it } from "vitest";

import {
  buildOcrPreviewRows,
  buildOcrWriteValues,
  canWriteOcrResult,
  extractOcrFieldNames,
  type OcrResult,
} from "./OCRTaskComposerPanel";

function baseResult(patch: Partial<OcrResult>): OcrResult {
  return {
    kind: "image",
    text: "",
    rows: [],
    fields: {},
    invoices: [],
    errors: [],
    ...patch,
  };
}

describe("OCRTaskComposerPanel helpers", () => {
  it("allows text-only OCR results to write without selected fields", () => {
    const result = baseResult({ text: "识别出的整段文本" });

    expect(extractOcrFieldNames(result)).toEqual([]);
    expect(canWriteOcrResult(result, [])).toBe(true);
    expect(buildOcrWriteValues(result, [])).toEqual([["识别出的整段文本"]]);
  });

  it("writes invoice rows by selected fields", () => {
    const result = baseResult({
      kind: "invoice",
      invoices: [
        { filename: "a.png", text: "", fields: { code: "001", amount: "12.30" }, rows: [] },
        { filename: "b.png", text: "", fields: { code: "002", amount: "45.60" }, rows: [] },
      ],
    });

    expect(extractOcrFieldNames(result)).toEqual(["code", "amount"]);
    expect(canWriteOcrResult(result, ["code", "amount"])).toBe(true);
    expect(buildOcrWriteValues(result, ["code", "amount"])).toEqual([
      ["code", "amount"],
      ["001", "12.30"],
      ["002", "45.60"],
    ]);
  });

  it("prefers invoice fields over top-level rows in invoice mode", () => {
    const result = baseResult({
      kind: "invoice",
      rows: [["表格列"], ["表格值"]],
      invoices: [
        { filename: "a.png", text: "", fields: { code: "001", amount: "12.30" }, rows: [] },
      ],
    });

    expect(extractOcrFieldNames(result)).toEqual(["code", "amount"]);
    expect(buildOcrWriteValues(result, ["code", "amount"])).toEqual([
      ["code", "amount"],
      ["001", "12.30"],
    ]);
  });

  it("previews invoice fields using the same rows that will be written", () => {
    const result = baseResult({
      kind: "invoice",
      rows: [["识别文本"], ["原始 OCR 文本"]],
      invoices: [
        { filename: "a.pdf", text: "", fields: { code: "001", amount: "12.30" }, rows: [] },
      ],
    });

    expect(buildOcrPreviewRows(result, ["code", "amount"])).toEqual([
      ["code", "amount"],
      ["001", "12.30"],
    ]);
  });
});
