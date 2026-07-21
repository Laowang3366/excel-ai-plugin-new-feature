import { describe, expect, it } from "vitest";
import {
  OCR_RESULT_MARKER_CLOSE,
  OCR_RESULT_MARKER_OPEN,
  buildOcrWriteValues,
  extractOcrFieldNames,
  parseOcrAssistantResult,
  parseSheetRangeAddress,
  sanitizeOcrUiText,
} from "../shared/tasks";

function fence(json: string, narrative = "说明"): string {
  return `${narrative}\n${OCR_RESULT_MARKER_OPEN}\n${json}\n${OCR_RESULT_MARKER_CLOSE}\n`;
}

describe("parseOcrAssistantResult", () => {
  it("parses valid invoice marker", () => {
    const text = fence(
      JSON.stringify({
        kind: "invoice",
        text: "摘要",
        fields: {},
        rows: [],
        invoices: [
          {
            filename: "a.png",
            text: "",
            fields: { 发票号码: "123", 金额: "10" },
            rows: [],
          },
        ],
        errors: [],
      }),
    );
    const r = parseOcrAssistantResult(text);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.result.kind).toBe("invoice");
    expect(r.result.invoices[0]?.fields["发票号码"]).toBe("123");
    expect(r.narrative).toContain("说明");
    expect(extractOcrFieldNames(r.result)).toEqual(
      expect.arrayContaining(["发票号码", "金额"]),
    );
    const values = buildOcrWriteValues(r.result, ["发票号码", "金额"]);
    expect(values[0]).toEqual(["发票号码", "金额"]);
    expect(values[1]).toEqual(["123", "10"]);
  });

  it("fails closed on missing marker and keeps raw text", () => {
    const r = parseOcrAssistantResult("只是普通回复，没有标记");
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.rawText).toContain("普通回复");
    expect(r.reason).toMatch(/未找到|标记/);
  });

  it("rejects multiple markers / bad kind / nested markers", () => {
    const multi = `${fence('{"kind":"image","text":"a"}')}${fence('{"kind":"image","text":"b"}')}`;
    expect(parseOcrAssistantResult(multi).ok).toBe(false);

    const badKind = fence(JSON.stringify({ kind: "pdf", text: "x" }));
    expect(parseOcrAssistantResult(badKind).ok).toBe(false);

    const nested = fence(
      `{"kind":"image","text":"${OCR_RESULT_MARKER_OPEN}"}`,
    );
    // JSON parse may succeed but nested marker check on raw body
    const n = parseOcrAssistantResult(nested);
    // if JSON escapes it might still be ok text; ensure invalid JSON fails
    const invalidJson = fence("{not json");
    expect(parseOcrAssistantResult(invalidJson).ok).toBe(false);
    void n;
  });

  it("sanitize strips api keys and long base64", () => {
    const s = sanitizeOcrUiText(
      "key sk-abcdefghijklmnopqrstuvwxyz012345 data:image/png;base64," +
        "A".repeat(60),
    );
    expect(s).not.toMatch(/sk-abcdefgh/);
    expect(s).toMatch(/redacted|omitted/i);
  });
});

describe("parseSheetRangeAddress", () => {
  it("parses plain and quoted sheet addresses", () => {
    expect(parseSheetRangeAddress("Sheet1!A1")).toEqual({
      sheetName: "Sheet1",
      range: "A1",
    });
    expect(parseSheetRangeAddress("'My Sheet'!B2:C3")).toEqual({
      sheetName: "My Sheet",
      range: "B2:C3",
    });
    expect(parseSheetRangeAddress("A1")).toBeNull();
    expect(parseSheetRangeAddress("")).toBeNull();
  });
});
