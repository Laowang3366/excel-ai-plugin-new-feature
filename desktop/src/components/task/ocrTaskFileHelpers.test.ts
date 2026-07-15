import { describe, expect, it } from "vitest";

import { isAcceptedOcrFile, isLikelyInvoiceFile, parseSheetRange } from "./ocrTaskFileHelpers";

function file(name: string, type = ""): File {
  return { name, type } as File;
}

describe("ocrTaskFileHelpers", () => {
  it("parses sheet-qualified ranges", () => {
    expect(parseSheetRange("Sheet1!A1")).toEqual({ sheetName: "Sheet1", range: "A1" });
    expect(parseSheetRange("'1月份出货明细'!D2:D142")).toEqual({
      sheetName: "1月份出货明细",
      range: "D2:D142",
    });
    expect(parseSheetRange("B2")).toEqual({ sheetName: "", range: "B2" });
  });

  it("accepts OCR image and PDF files by mime type or extension", () => {
    expect(isAcceptedOcrFile(file("scan.bin", "image/png"))).toBe(true);
    expect(isAcceptedOcrFile(file("invoice.PDF"))).toBe(true);
    expect(isAcceptedOcrFile(file("notes.txt", "text/plain"))).toBe(false);
  });

  it("detects likely invoice files from common names", () => {
    expect(isLikelyInvoiceFile(file("7月发票.png"))).toBe(true);
    expect(isLikelyInvoiceFile(file("invoice-001.jpg"))).toBe(true);
    expect(isLikelyInvoiceFile(file("scan.png"))).toBe(false);
  });
});
