/**
 * Destination planning: dedicated Pivots sheet + next free address (desktop parity).
 */
import { parseA1Cell, toA1 } from "./a1Address";
import { validateBareA1 } from "./officeJsChartSource";
import type { ExcelRange, ExcelRequestContextWithPivot, ExcelWorksheetWithPivot } from "./officeJsPivotTypes";
import { PIVOT_DEFAULT_SHEET } from "./pivotTypes";

export type PivotDestinationPlan = {
  useDedicatedSheet: boolean;
  sheetName: string | null;
  address: string;
};

/**
 * Empty → Pivots sheet auto. Sheet!A1 or bare A1 (caller binds bare to source sheet).
 * Rejects external/3D/multi-area/structured refs.
 */
export function parsePivotDestination(destination: string | undefined | null): PivotDestinationPlan {
  if (destination == null || String(destination).trim() === "") {
    return { useDedicatedSheet: true, sheetName: null, address: "A1" };
  }
  const raw = String(destination).trim();
  if (raw.includes("[") || raw.includes("]")) {
    throw new Error("destination structured/external references are not supported");
  }
  if (raw.includes(",")) {
    throw new Error("destination multi-area is not supported");
  }
  // 3D: Sheet1:Sheet2!A1
  if (/!/.test(raw) && /:/.test(raw.split("!")[0] ?? "")) {
    throw new Error("destination 3D references are not supported");
  }

  let sheetName: string | null = null;
  let a1Part: string;
  if (raw.includes("!")) {
    const bang = raw.indexOf("!");
    let sheetPart = raw.slice(0, bang);
    a1Part = raw.slice(bang + 1);
    if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
      sheetPart = sheetPart.slice(1, -1).replace(/''/g, "'");
    }
    if (sheetPart.trim() === "") throw new Error("destination sheet name must be non-empty");
    if (/[\[\]:*?/\\]/.test(sheetPart)) {
      throw new Error("destination sheet name contains illegal characters");
    }
    sheetName = sheetPart;
  } else {
    a1Part = raw;
  }
  const address = validateBareA1(a1Part, "destination");
  return { useDedicatedSheet: false, sheetName, address };
}

export async function ensurePivotSheet(
  context: ExcelRequestContextWithPivot,
): Promise<ExcelWorksheetWithPivot> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const items = sheets.items ?? [];
  for (const sheet of items) {
    if (String(sheet.name).toLowerCase() === PIVOT_DEFAULT_SHEET.toLowerCase()) {
      return sheet as ExcelWorksheetWithPivot;
    }
  }
  const created = sheets.add(PIVOT_DEFAULT_SHEET) as ExcelWorksheetWithPivot;
  created.load("name");
  await context.sync();
  return created;
}

/**
 * Next free top-left under existing pivots on the sheet (A1 or A{lastBottom+3}).
 * Uses PivotLayout.getRange() (ExcelApi 1.8).
 */
export async function nextPivotDestinationAddress(
  context: ExcelRequestContextWithPivot,
  sheet: ExcelWorksheetWithPivot,
): Promise<string> {
  const pivots = sheet.pivotTables;
  pivots.load("items/name");
  await context.sync();
  const items = pivots.items ?? [];
  if (items.length === 0) return "A1";

  let lastRow = 0;
  for (const pivot of items) {
    const range = pivot.layout.getRange();
    range.load("address,rowCount,rowIndex");
    // Prefer rowIndex+rowCount when available; fall back to parsing address.
  }
  await context.sync();

  for (const pivot of items) {
    const range = pivot.layout.getRange() as ExcelRange & {
      address?: string;
      rowCount?: number;
      rowIndex?: number;
    };
    if (typeof range.rowIndex === "number" && typeof range.rowCount === "number") {
      const bottom = range.rowIndex + range.rowCount; // 0-based end exclusive → 1-based bottom = rowIndex+rowCount
      lastRow = Math.max(lastRow, bottom);
      continue;
    }
    const address = String(range.address ?? "");
    const bare = address.includes("!") ? address.split("!")[1]! : address;
    const parts = bare.split(":");
    const end = parts[parts.length - 1] ?? bare;
    const cell = parseA1Cell(end);
    if (cell) lastRow = Math.max(lastRow, cell.row + 1);
  }
  return lastRow === 0 ? "A1" : toA1(lastRow + 2, 0); // lastRow is 1-based bottom; +2 → +3 from 1-based inclusive
}

export function formatSheetA1(sheetName: string, bareA1: string): string {
  const needsQuotes = /[\s'!]/.test(sheetName) || !/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheetName);
  if (needsQuotes) return `'${sheetName.replace(/'/g, "''")}'!${bareA1}`;
  return `${sheetName}!${bareA1}`;
}
