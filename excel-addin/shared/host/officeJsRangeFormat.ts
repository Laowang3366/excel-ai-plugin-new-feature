/**
 * Range format load/read helpers for Office.js adapters.
 */
import type { ExcelRange } from "./officeJsExcelTypes";
import type { RangeFormat } from "./types";
import { firstNumberFormat } from "./officeJsNormalize";

export function loadRangeFormat(range: ExcelRange): void {
  range.load("address,numberFormat");
  range.format.load("horizontalAlignment,verticalAlignment,wrapText");
  range.format.font.load("name,size,bold,color");
  range.format.fill.load("color");
}

export function readFormatFromRange(range: ExcelRange): RangeFormat {
  return {
    fontName: range.format.font.name ?? null,
    fontSize: range.format.font.size ?? null,
    fontBold: range.format.font.bold ?? null,
    fontColor: range.format.font.color ?? null,
    fillColor: range.format.fill.color ?? null,
    numberFormat: firstNumberFormat(range.numberFormat),
    horizontalAlignment: range.format.horizontalAlignment ?? null,
    verticalAlignment: range.format.verticalAlignment ?? null,
    wrapText: range.format.wrapText ?? null,
  };
}
