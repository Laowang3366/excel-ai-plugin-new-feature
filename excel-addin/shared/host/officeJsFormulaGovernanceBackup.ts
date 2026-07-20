/**
 * WENGGE_FORMULA_BACKUP_V1 workbook sheet materialization (Office.js).
 */
import {
  FORMULA_BACKUP_HEADERS,
  FORMULA_BACKUP_MAGIC,
  createBackupRows,
  decodeBackupSheet,
  encodeBackupSheet,
  summarizeBackups,
  type FormulaBackupRow,
  type FormulaCellRecord,
} from "../formulaGovernance";
import {
  FORMULA_BACKUP_SHEET_PREFIX,
  type FormulaBackupsInspectInfo,
} from "./formulaGovernanceTypes";
import { bareAddress, isBackupSheetName } from "./officeJsFormulaGovernanceCollect";
import type { ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";

export function newBackupId(): string {
  // Prefer crypto.randomUUID when available; fall back to time+random.
  const c = globalThis.crypto as Crypto | undefined;
  if (c && typeof c.randomUUID === "function") {
    return c.randomUUID().replace(/-/g, "");
  }
  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2, 10)}`;
}

export async function findBackupSheet(
  context: ExcelRequestContext,
  create: boolean,
): Promise<{ sheet: ExcelWorksheet | null; created: boolean; limitations: string[] }> {
  const limitations: string[] = [];
  context.workbook.worksheets.load("items/name");
  await context.sync();

  for (const ws of context.workbook.worksheets.items) {
    ws.load("name");
    await context.sync();
    if (!isBackupSheetName(ws.name)) continue;
    const a1 = ws.getRange("A1");
    a1.load("values");
    await context.sync();
    const magic = String(a1.values?.[0]?.[0] ?? "").trim();
    if (magic === FORMULA_BACKUP_MAGIC) {
      return { sheet: ws, created: false, limitations };
    }
    limitations.push(
      `sheet ${ws.name} has backup prefix but invalid magic; left untouched (no overwrite)`,
    );
  }

  if (!create) return { sheet: null, created: false, limitations };

  const stamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const name = `${FORMULA_BACKUP_SHEET_PREFIX}${stamp}`.slice(0, 31);
  const sheet = context.workbook.worksheets.add(name);
  sheet.getRange("A1").values = [[FORMULA_BACKUP_MAGIC]];
  sheet.getRange("A2").values = [[...FORMULA_BACKUP_HEADERS]];
  // Expand header row across columns
  const headerRange = sheet.getRange("A2:J2");
  headerRange.values = [[...FORMULA_BACKUP_HEADERS]];
  sheet.visibility = "VeryHidden";
  sheet.load("name");
  await context.sync();
  limitations.push(`created backup sheet ${sheet.name} (VeryHidden)`);
  return { sheet, created: true, limitations };
}

export async function readBackupMatrix(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
): Promise<unknown[][]> {
  const used = sheet.getUsedRangeOrNullObject(false);
  used.load("isNullObject,address,values,rowCount,columnCount");
  await context.sync();
  if (used.isNullObject) {
    return [[FORMULA_BACKUP_MAGIC], [...FORMULA_BACKUP_HEADERS]];
  }
  const values = used.values as unknown[][];
  return Array.isArray(values) ? values : [];
}

export async function lastBackupDataRow(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
): Promise<number> {
  // Excel rows are 1-based; data starts at row 3. Return last used row index (1-based).
  const used = sheet.getUsedRangeOrNullObject(false);
  used.load("isNullObject,rowCount,address");
  await context.sync();
  if (used.isNullObject) return 2;
  const address = typeof used.address === "string" ? used.address : "A1";
  const bare = bareAddress(address);
  const m = /:(\$?[A-Z]+\$?)?(\d+)$/i.exec(bare) || /^[A-Z]+(\d+)$/i.exec(bare);
  if (m) {
    const row = Number(m[m.length - 1]);
    return Number.isFinite(row) ? Math.max(2, row) : 2;
  }
  const rc = typeof used.rowCount === "number" ? used.rowCount : 2;
  return Math.max(2, rc);
}

export async function appendBackupRows(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rows: FormulaBackupRow[],
): Promise<void> {
  if (rows.length === 0) return;
  const last = await lastBackupDataRow(context, sheet);
  const startRow = last + 1; // 1-based
  const grid = rows.map((row) => [
    row.backupId,
    row.createdAt,
    row.sheet,
    row.address,
    row.formula,
    row.formulaR1C1,
    row.numberFormat,
    row.locked ? "1" : "0",
    row.spillAddress,
    row.sourceRange,
  ]);
  const endRow = startRow + grid.length - 1;
  const range = sheet.getRange(`A${startRow}:J${endRow}`);
  range.values = grid;
  sheet.visibility = "VeryHidden";
  await context.sync();
}

export async function writeFormulaBackup(
  context: ExcelRequestContext,
  cells: FormulaCellRecord[],
  options: { backupId: string; sourceRange: string; create: boolean },
): Promise<{ backupId: string; sheetName: string; rowCount: number; limitations: string[] }> {
  const { sheet, limitations } = await findBackupSheet(context, options.create);
  if (!sheet) {
    throw new Error("formula_backup_sheet_unavailable");
  }
  const rows = createBackupRows(cells, {
    backupId: options.backupId,
    sourceRange: options.sourceRange,
  });
  if (rows.length === 0 && cells.some((c) => c.formula.startsWith("="))) {
    // createBackupRows skips non-formulas; empty is ok for no-op convert
  }
  await appendBackupRows(context, sheet, rows);
  sheet.load("name");
  await context.sync();
  return {
    backupId: options.backupId,
    sheetName: sheet.name,
    rowCount: rows.length,
    limitations,
  };
}

export async function inspectBackupSheet(
  context: ExcelRequestContext,
): Promise<FormulaBackupsInspectInfo> {
  const limitations: string[] = [];
  const found = await findBackupSheet(context, false);
  limitations.push(...found.limitations);
  if (!found.sheet) {
    return {
      backups: [],
      backupCount: 0,
      backupSheetName: null,
      skippedRows: [],
      limitations,
    };
  }
  found.sheet.load("name");
  await context.sync();
  const matrix = await readBackupMatrix(context, found.sheet);
  const decoded = decodeBackupSheet(matrix);
  if (!decoded.ok && !decoded.grid) {
    return {
      backups: [],
      backupCount: 0,
      backupSheetName: found.sheet.name,
      skippedRows: decoded.skipped,
      limitations: [...limitations, decoded.error ?? "invalid backup sheet"],
      headerError: decoded.error,
    };
  }
  const rows = decoded.grid?.rows ?? [];
  const backups = summarizeBackups(rows);
  if (decoded.error) limitations.push(decoded.error);
  if (decoded.skipped.length > 0) {
    limitations.push(`skipped ${decoded.skipped.length} corrupt/incomplete backup row(s)`);
  }
  return {
    backups,
    backupCount: backups.length,
    backupSheetName: found.sheet.name,
    skippedRows: decoded.skipped,
    limitations,
    headerError: decoded.error,
  };
}

export async function loadBackupRows(
  context: ExcelRequestContext,
): Promise<{
  sheet: ExcelWorksheet | null;
  rows: FormulaBackupRow[];
  skipped: number[];
  error?: string;
  limitations: string[];
}> {
  const found = await findBackupSheet(context, false);
  if (!found.sheet) {
    return {
      sheet: null,
      rows: [],
      skipped: [],
      error: "formula_backup_not_found",
      limitations: found.limitations,
    };
  }
  const matrix = await readBackupMatrix(context, found.sheet);
  const decoded = decodeBackupSheet(matrix);
  if (!decoded.ok && !decoded.grid) {
    return {
      sheet: found.sheet,
      rows: [],
      skipped: decoded.skipped,
      error: decoded.error ?? "formula_backup_corrupt",
      limitations: found.limitations,
    };
  }
  return {
    sheet: found.sheet,
    rows: decoded.grid?.rows ?? [],
    skipped: decoded.skipped,
    error: decoded.error,
    limitations: found.limitations,
  };
}

/** Re-encode full grid (for removeAfterRestore). */
export async function rewriteBackupSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rows: FormulaBackupRow[],
): Promise<void> {
  const grid = encodeBackupSheet(rows);
  const used = sheet.getUsedRangeOrNullObject(false);
  used.load("isNullObject");
  await context.sync();
  if (!used.isNullObject) {
    used.clear();
    await context.sync();
  }
  if (grid.length === 0) return;
  const cols = 10;
  const padded = grid.map((row) => {
    const out = row.map((v) => (v == null ? "" : String(v)));
    while (out.length < cols) out.push("");
    return out.slice(0, cols);
  });
  const range = sheet.getRange(`A1:J${padded.length}`);
  range.values = padded;
  sheet.visibility = "VeryHidden";
  await context.sync();
}
