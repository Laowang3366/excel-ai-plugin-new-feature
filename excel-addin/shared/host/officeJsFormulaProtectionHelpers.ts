/**
 * Formula protection via Range.formulas + Range.format.protection.locked (ExcelApi 1.2)
 * and Worksheet.protection (existing). Does not whole-sheet lock as a substitute for
 * per-formula locking.
 */
import type {
  FormulaProtectionCommand,
  FormulaProtectionInspectInfo,
  FormulaProtectionInspectInput,
  FormulaProtectionScope,
  FormulaProtectionSheetSummary,
} from "./formulaProtectionTypes";
import type {
  ExcelRange,
  ExcelRequestContext,
  ExcelWorksheet,
} from "./officeJsRuntime";

export const LOCKED_EVIDENCE =
  "Range.format.protection.locked requires ExcelApi 1.2 to lock/unlock formula cells without whole-sheet-only fake protection";

export const MAX_FORMULA_CELLS = 5_000;
export const LOAD_BATCH = 80;

export function isExcelApi12Supported(): boolean {
  const office = (
    globalThis as unknown as {
      Office?: {
        context?: {
          requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
        };
      };
    }
  ).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.2");
  } catch {
    return false;
  }
}

export function requireScope(scope: FormulaProtectionScope, sheetName?: string, range?: string): void {
  if (scope === "workbook") return;
  if (!sheetName || sheetName.trim() === "") {
    throw new Error("sheetName is required for scope sheet|target");
  }
  if (scope === "target" && (!range || range.trim() === "")) {
    throw new Error("range is required for scope=target");
  }
}

export function isFormulaCell(value: unknown): boolean {
  return typeof value === "string" && value.trim().startsWith("=");
}

export type FormulaCoord = { row: number; col: number };

export function collectFormulaCoords(formulas: unknown[][]): FormulaCoord[] {
  const coords: FormulaCoord[] = [];
  for (let r = 0; r < formulas.length; r += 1) {
    const row = formulas[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      if (isFormulaCell(row[c])) coords.push({ row: r, col: c });
    }
  }
  return coords;
}

export async function resolveRange(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress?: string,
): Promise<ExcelRange> {
  if (rangeAddress && rangeAddress.trim()) {
    const bare = rangeAddress.includes("!")
      ? rangeAddress.slice(rangeAddress.lastIndexOf("!") + 1)
      : rangeAddress;
    return sheet.getRange(bare);
  }
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("isNullObject,address");
  await context.sync();
  if (used.isNullObject) {
    // Empty sheet — use A1 as empty scan window.
    return sheet.getRange("A1");
  }
  // Re-bind via address so getCell/formulas/protection work on a real Range proxy.
  const address = typeof used.address === "string" ? used.address : "A1";
  const bare = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  return sheet.getRange(bare || "A1");
}

export async function inspectSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress: string | undefined,
  limitations: string[],
): Promise<FormulaProtectionSheetSummary> {
  sheet.load("name");
  sheet.protection.load("protected");
  const range = await resolveRange(context, sheet, rangeAddress);
  range.load("address,formulas,rowCount,columnCount");
  await context.sync();

  const sheetName = typeof sheet.name === "string" ? sheet.name : "";
  const address = typeof range.address === "string" ? range.address : "";
  const formulas = Array.isArray(range.formulas) ? (range.formulas as unknown[][]) : [];
  let coords = collectFormulaCoords(formulas);
  const sheetLimitations: string[] = [];

  if (coords.length > MAX_FORMULA_CELLS) {
    sheetLimitations.push(
      `formula cell cap ${MAX_FORMULA_CELLS}: scanned first ${MAX_FORMULA_CELLS} of ${coords.length}`,
    );
    limitations.push(
      `${sheetName}: formula cell cap ${MAX_FORMULA_CELLS} (found ${coords.length})`,
    );
    coords = coords.slice(0, MAX_FORMULA_CELLS);
  }

  let lockedFormulaCount = 0;
  // Probe protection API once on anchor cell.
  if (coords.length > 0) {
    const probe = range.getCell(coords[0]!.row, coords[0]!.col);
    const protection = probe.format?.protection;
    if (protection == null || typeof protection.load !== "function") {
      throw new Error("Range.format.protection.locked missing (ExcelApi 1.2 required)");
    }
  }

  for (let i = 0; i < coords.length; i += LOAD_BATCH) {
    const batch = coords.slice(i, i + LOAD_BATCH);
    const cells = batch.map((coord) => {
      const cell = range.getCell(coord.row, coord.col);
      cell.format.protection.load("locked");
      return cell;
    });
    await context.sync();
    for (const cell of cells) {
      if (cell.format.protection.locked === true) lockedFormulaCount += 1;
    }
  }

  return {
    sheetName,
    address,
    formulaCount: coords.length,
    lockedFormulaCount,
    sheetProtected: sheet.protection.protected === true,
    limitations: sheetLimitations,
  };
}

export async function inspectAll(
  context: ExcelRequestContext,
  input: FormulaProtectionInspectInput,
): Promise<FormulaProtectionInspectInfo> {
  requireScope(input.scope, input.sheetName, input.range);
  const limitations: string[] = [];
  const sheets: FormulaProtectionSheetSummary[] = [];

  if (input.scope === "workbook") {
    context.workbook.worksheets.load("items/name");
    await context.sync();
    for (const ws of context.workbook.worksheets.items) {
      sheets.push(await inspectSheet(context, ws, undefined, limitations));
    }
  } else {
    const sheet = context.workbook.worksheets.getItem(input.sheetName!);
    sheets.push(await inspectSheet(context, sheet, input.range, limitations));
  }

  const formulaCount = sheets.reduce((sum, s) => sum + s.formulaCount, 0);
  const lockedFormulaCount = sheets.reduce((sum, s) => sum + s.lockedFormulaCount, 0);
  return {
    scope: input.scope,
    sheets,
    formulaCount,
    lockedFormulaCount,
    limitations,
  };
}

export async function setFormulaLocks(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress: string | undefined,
  locked: boolean,
  unlockInputs: boolean,
  limitations: string[],
): Promise<{ formulaCount: number }> {
  sheet.load("name");
  const range = await resolveRange(context, sheet, rangeAddress);
  range.load("address,formulas,rowCount,columnCount");
  await context.sync();

  const formulas = Array.isArray(range.formulas) ? (range.formulas as unknown[][]) : [];
  let coords = collectFormulaCoords(formulas);
  if (coords.length > MAX_FORMULA_CELLS) {
    limitations.push(
      `formula cell cap ${MAX_FORMULA_CELLS}: modified first ${MAX_FORMULA_CELLS} of ${coords.length}`,
    );
    coords = coords.slice(0, MAX_FORMULA_CELLS);
  }

  if (locked && unlockInputs) {
    // Unlock whole target range so non-formula inputs stay editable under sheet protect.
    // Scope is strictly this range — never whole workbook.
    if (range.format?.protection == null || typeof range.format.protection.load !== "function") {
      throw new Error("Range.format.protection.locked missing (ExcelApi 1.2 required)");
    }
    range.format.protection.locked = false;
    await context.sync();
    limitations.push(
      "unlockInputs: unlocked all cells in target range before locking formula cells (inputs outside range unchanged)",
    );
  }

  for (let i = 0; i < coords.length; i += LOAD_BATCH) {
    const batch = coords.slice(i, i + LOAD_BATCH);
    for (const coord of batch) {
      const cell = range.getCell(coord.row, coord.col);
      if (cell.format?.protection == null) {
        throw new Error("Range.format.protection.locked missing (ExcelApi 1.2 required)");
      }
      cell.format.protection.locked = locked;
    }
    await context.sync();
  }

  return { formulaCount: coords.length };
}

export function verifyManage(
  command: FormulaProtectionCommand,
  unlockInputs: boolean,
  protectSheet: boolean,
  before: FormulaProtectionInspectInfo,
  after: FormulaProtectionInspectInfo,
): { verified: boolean; limitations: string[] } {
  const limitations: string[] = [];
  if (after.formulaCount === 0) {
    limitations.push("no formula cells in scope; lock/unlock was a no-op on cells");
    return { verified: true, limitations };
  }
  if (command === "lock") {
    if (after.lockedFormulaCount !== after.formulaCount) {
      limitations.push(
        `lock verification failed: lockedFormulaCount=${after.lockedFormulaCount} formulaCount=${after.formulaCount}`,
      );
      return { verified: false, limitations };
    }
    if (protectSheet) {
      const unprotected = after.sheets.filter((s) => !s.sheetProtected);
      if (unprotected.length > 0) {
        limitations.push(
          `protectSheet verification failed on: ${unprotected.map((s) => s.sheetName).join(",")}`,
        );
        return { verified: false, limitations };
      }
    }
  } else {
    if (after.lockedFormulaCount !== 0) {
      limitations.push(
        `unlock verification failed: lockedFormulaCount=${after.lockedFormulaCount} (expected 0)`,
      );
      return { verified: false, limitations };
    }
  }
  if (before.formulaCount !== after.formulaCount) {
    limitations.push(
      `formulaCount changed during manage (${before.formulaCount}→${after.formulaCount}); using post-write counts`,
    );
  }
  if (unlockInputs && command === "lock") {
    limitations.push(
      "unlockInputs applied only within target range; input cells outside target remain as-is",
    );
  }
  return { verified: true, limitations };
}

