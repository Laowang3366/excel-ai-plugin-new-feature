/**
 * Destination planning: dedicated Pivots sheet + next free address (desktop parity).
 * Desktop NextPivotDestinationAddress: A1 or A{lastBottom + 3} (1-based bottom of TableRange2).
 */
import { parseA1Cell } from "./a1Address";
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

/**
 * Pure desktop-aligned placement: lastBottom is 1-based inclusive bottom row of existing pivots.
 * Desktop: lastRow == 0 ? "A1" : $"A{lastRow + 3}".
 */
export function computeNextPivotAddress(lastBottom1Based: number): string {
  if (!Number.isFinite(lastBottom1Based) || lastBottom1Based <= 0) return "A1";
  return `A${Math.floor(lastBottom1Based) + 3}`;
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
 * Next free top-left under existing pivots on the sheet (desktop lastBottom+3).
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

  for (const pivot of items) {
    const range = pivot.layout.getRange();
    range.load("address,rowCount,rowIndex");
  }
  await context.sync();

  let lastBottom = 0;
  for (const pivot of items) {
    const range = pivot.layout.getRange() as ExcelRange & {
      address?: string;
      rowCount?: number;
      rowIndex?: number;
    };
    if (typeof range.rowIndex === "number" && typeof range.rowCount === "number") {
      // Excel.Range.rowIndex is 0-based; 1-based bottom = rowIndex + rowCount
      const bottom = range.rowIndex + range.rowCount;
      lastBottom = Math.max(lastBottom, bottom);
      continue;
    }
    const address = String(range.address ?? "");
    const bare = address.includes("!") ? address.split("!")[1]! : address;
    const parts = bare.split(":");
    const end = parts[parts.length - 1] ?? bare;
    const cell = parseA1Cell(end);
    if (cell) lastBottom = Math.max(lastBottom, cell.row + 1);
  }
  return computeNextPivotAddress(lastBottom);
}

export function formatSheetA1(sheetName: string, bareA1: string): string {
  const needsQuotes = /[\s'!]/.test(sheetName) || !/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheetName);
  if (needsQuotes) return `'${sheetName.replace(/'/g, "''")}'!${bareA1}`;
  return `${sheetName}!${bareA1}`;
}
