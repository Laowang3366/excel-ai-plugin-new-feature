/** Helpers for workbook.template.apply (empty detection, targets, freeze verify). */
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";
import type { WorkbookTemplateApplyInput } from "./workbookTemplateTypes";
import {
  requireBoolean,
  requireNonEmptyString,
  requireNonNegativeInt,
  requireParseableA1Range,
  requirePositiveInt,
} from "./officeJsTemplateReadback";

/**
 * Empty detection without bulk text: dims first; text/values only for 1×1.
 */
export async function isEmptyUsedRange(
  used: ExcelRange & { isNullObject?: unknown },
  context: ExcelRequestContext,
): Promise<boolean> {
  const isNull = requireBoolean(used.isNullObject, "UsedRange.isNullObject");
  if (isNull) return true;
  used.load("address,rowCount,columnCount");
  await context.sync();
  const rows = requirePositiveInt(used.rowCount, "UsedRange.rowCount");
  const cols = requirePositiveInt(used.columnCount, "UsedRange.columnCount");
  if (rows !== 1 || cols !== 1) return false;
  used.load("text");
  await context.sync();
  const text = (used as ExcelRange & { text?: unknown }).text;
  if (text === null || text === undefined || text === "") return true;
  if (typeof text === "string" && text.trim() === "") return true;
  if (typeof text !== "string") {
    used.load("values");
    await context.sync();
    const v = used.values?.[0]?.[0];
    if (v == null || (typeof v === "string" && v.trim() === "")) return true;
  }
  return false;
}

export async function resolveApplyTargets(
  context: ExcelRequestContext,
  input: WorkbookTemplateApplyInput,
): Promise<ExcelWorksheet[]> {
  if (input.sheetNames !== undefined && input.sheetNames.length === 0) {
    throw new Error("sheetNames must not be an empty array");
  }
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  const active = context.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  await context.sync();

  const items = sheets.items;
  if (!Array.isArray(items)) throw new Error("WorksheetCollection.items is not an array");
  if (items.length > 500) {
    throw new Error("workbook exceeds 500 worksheets (resource-limit)");
  }

  const byLower = new Map<string, ExcelWorksheet>();
  for (const sheet of items) {
    const name = requireNonEmptyString(sheet.name, "Worksheet.name");
    byLower.set(name.toLowerCase(), sheet);
  }

  let selected: ExcelWorksheet[] =
    input.sheetNames && input.sheetNames.length > 0
      ? input.sheetNames.map((requested) => {
          const hit = byLower.get(requested.toLowerCase());
          if (!hit) throw new Error(`sheet not found: ${requested}`);
          return hit;
        })
      : items.slice();

  if (!input.allSheets) {
    const activeName = requireNonEmptyString(active.name, "ActiveWorksheet.name");
    selected = selected.filter(
      (s) =>
        requireNonEmptyString(s.name, "Worksheet.name").toLowerCase() ===
        activeName.toLowerCase(),
    );
  }

  const seen = new Set<string>();
  const unique: ExcelWorksheet[] = [];
  for (const sheet of selected) {
    const name = requireNonEmptyString(sheet.name, "Worksheet.name");
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(sheet);
  }
  if (unique.length === 0) {
    throw new Error("no target worksheets after sheetNames/allSheets selection");
  }
  return unique;
}

export function verifyFreezeReadback(
  loc: ExcelRange & { isNullObject?: unknown; address?: unknown },
  freezeRows: number,
): number {
  const locNull = requireBoolean(loc.isNullObject, "freeze.isNullObject");
  if (locNull) {
    if (freezeRows !== 0) throw new Error("freeze location is null but freezeRows > 0");
    return 0;
  }
  const freezeRowCount = requireNonNegativeInt(loc.rowCount, "freeze.rowCount");
  requireNonNegativeInt(loc.columnCount, "freeze.columnCount");
  requireParseableA1Range(loc.address, "freeze.address");
  if (freezeRowCount !== freezeRows) {
    throw new Error(`freeze rowCount readback mismatch: ${freezeRowCount} !== ${freezeRows}`);
  }
  return freezeRowCount;
}
