/**
 * WENGGE_FORMULA_BACKUP_V1 workbook sheet materialization (Office.js).
 * Formula columns are written as text (numberFormat "@" + apostrophe encode).
 */
import {
  FORMULA_BACKUP_HEADERS,
  FORMULA_BACKUP_MAGIC,
  createBackupRows,
  decodeBackupLiteral,
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
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";

export function newBackupId(): string {
  const c = globalThis.crypto as Crypto | undefined;
  if (c && typeof c.randomUUID === "function") {
    return c.randomUUID().replace(/-/g, "");
  }
  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2, 10)}`;
}

/** Write a values matrix only when dimensions match; set text format first. */
export async function writeTextMatrix(
  context: ExcelRequestContext,
  range: ExcelRange,
  matrix: string[][],
): Promise<void> {
  const rows = matrix.length;
  const cols = rows > 0 ? matrix[0]!.length : 0;
  for (const row of matrix) {
    if (row.length !== cols) {
      throw new Error(`backup matrix jagged: expected ${cols} cols`);
    }
  }
  // numberFormat "@" reduces evaluation of formula-like strings on values write
  range.numberFormat = Array.from({ length: rows }, () => Array(cols).fill("@"));
  range.values = matrix;
  await context.sync();
}

async function setVeryHidden(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
): Promise<void> {
  sheet.visibility = "VeryHidden";
  sheet.load("visibility,name");
  await context.sync();
  const v = String(sheet.visibility ?? "").toLowerCase();
  if (!v.includes("veryhidden")) {
    throw new Error(
      `backup_sheet_visibility_unavailable: expected VeryHidden, got ${sheet.visibility ?? "(empty)"}`,
    );
  }
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
    const magic = decodeBackupLiteral(String(a1.values?.[0]?.[0] ?? "")).trim();
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
  // Hide before writing any formula-bearing content
  await setVeryHidden(context, sheet);
  await writeTextMatrix(context, sheet.getRange("A1"), [[FORMULA_BACKUP_MAGIC]]);
  // Headers: single write A2:J2 only (never A2 1x10 into a 1x1 range)
  await writeTextMatrix(context, sheet.getRange("A2:J2"), [[...FORMULA_BACKUP_HEADERS]]);
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
  const startRow = last + 1;
  // encodeBackupSheet includes magic+headers; data-only grid from rows:
  const full = encodeBackupSheet(rows);
  const data = full.slice(2); // data rows only, already literal-encoded
  const endRow = startRow + data.length - 1;
  const range = sheet.getRange(`A${startRow}:J${endRow}`);
  await writeTextMatrix(context, range, data);
  await setVeryHidden(context, sheet);
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
  await setVeryHidden(context, sheet);
  const rows = createBackupRows(cells, {
    backupId: options.backupId,
    sourceRange: options.sourceRange,
  });
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

/**
 * Strict protocol check for restore (fail closed).
 * Inspect may still report partial parses via decodeBackupSheet.
 */
export function strictDecodeBackup(matrix: unknown[][]): {
  ok: true;
  rows: FormulaBackupRow[];
  skipped: number[];
} | {
  ok: false;
  error: string;
  skipped: number[];
} {
  if (!matrix.length) {
    return { ok: false, error: "formula_backup_corrupt: empty backup sheet", skipped: [] };
  }
  const magic = decodeBackupLiteral(String(matrix[0]?.[0] ?? "")).trim();
  if (magic !== FORMULA_BACKUP_MAGIC) {
    return {
      ok: false,
      error: `formula_backup_corrupt: invalid magic: ${magic || "(empty)"}`,
      skipped: [],
    };
  }
  const headerRow = matrix[1] ?? [];
  const headerOk = FORMULA_BACKUP_HEADERS.every(
    (h, i) => String(headerRow[i] ?? "").trim() === h,
  );
  if (!headerOk) {
    return {
      ok: false,
      error: "formula_backup_corrupt: header mismatch",
      skipped: [],
    };
  }
  const decoded = decodeBackupSheet(matrix);
  if (!decoded.ok || !decoded.grid) {
    return {
      ok: false,
      error: decoded.error ?? "formula_backup_corrupt",
      skipped: decoded.skipped,
    };
  }
  if (decoded.error) {
    return {
      ok: false,
      error: `formula_backup_corrupt: ${decoded.error}`,
      skipped: decoded.skipped,
    };
  }
  // Restore fail-closed: any corrupt/incomplete data row aborts (inspect may still report skipped).
  if (decoded.skipped.length > 0) {
    return {
      ok: false,
      error: `formula_backup_corrupt: skipped data row(s) ${decoded.skipped.join(",")}`,
      skipped: decoded.skipped,
    };
  }
  return { ok: true, rows: decoded.grid.rows, skipped: [] };
}


/** Clear used area and rewrite magic+headers+remaining rows (removeAfterRestore). */
export async function rewriteBackupSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rows: FormulaBackupRow[],
): Promise<void> {
  const used = sheet.getUsedRangeOrNullObject(false);
  used.load("isNullObject");
  await context.sync();
  if (!used.isNullObject) {
    used.clear();
    await context.sync();
  }
  const grid = encodeBackupSheet(rows);
  // Pad to 10 columns so A1:J{n} is rectangular (magic row is 1-wide in encode).
  const cols = 10;
  const padded = grid.map((row) => {
    const out = row.map((v) => (v == null ? "" : String(v)));
    while (out.length < cols) out.push("");
    return out.slice(0, cols);
  });
  if (padded.length === 0) {
    await writeTextMatrix(context, sheet.getRange("A1"), [[FORMULA_BACKUP_MAGIC]]);
    await writeTextMatrix(context, sheet.getRange("A2:J2"), [[...FORMULA_BACKUP_HEADERS]]);
  } else {
    await writeTextMatrix(context, sheet.getRange(`A1:J${padded.length}`), padded);
  }
  await setVeryHidden(context, sheet);
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

export async function loadBackupRowsStrict(
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
    const corruptHint = found.limitations.some((l) => /invalid magic/i.test(l));
    return {
      sheet: null,
      rows: [],
      skipped: [],
      error: corruptHint
        ? "formula_backup_corrupt: invalid magic on backup-prefixed sheet(s)"
        : "formula_backup_not_found",
      limitations: found.limitations,
    };
  }
  const matrix = await readBackupMatrix(context, found.sheet);
  const strict = strictDecodeBackup(matrix);
  if (!strict.ok) {
    return {
      sheet: found.sheet,
      rows: [],
      skipped: strict.skipped,
      error: strict.error,
      limitations: found.limitations,
    };
  }
  return {
    sheet: found.sheet,
    rows: strict.rows,
    skipped: strict.skipped,
    limitations: found.limitations,
  };
}
