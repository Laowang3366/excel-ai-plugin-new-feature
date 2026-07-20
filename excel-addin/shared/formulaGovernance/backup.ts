/**
 * WENGGE_FORMULA_BACKUP_V1 encode/decode (desktop GetBackupSheet / WriteBackup).
 * Rows are workbook-sheet materializations — never session-only storage.
 */

import {
  FORMULA_BACKUP_HEADERS,
  FORMULA_BACKUP_MAGIC,
  type FormulaBackupRow,
  type FormulaBackupSummary,
  type FormulaCellRecord,
  type FormulaRestorePlan,
  type FormulaRestorePlanItem,
} from "./types";
import { normalizeA1Address } from "./address";


/**
 * Persist formula-like text without host evaluation.
 * Leading apostrophe is the classic Excel "store as text" marker; paired with
 * host numberFormat "@" when writing values/Value2.
 */
export function encodeBackupLiteral(value: string): string {
  if (!value) return value;
  const ch = value[0];
  if (ch === "=" || ch === "+" || ch === "-" || ch === "@") {
    return `'${value}`;
  }
  return value;
}

/** Inverse of encodeBackupLiteral; host may already strip the apostrophe. */
export function decodeBackupLiteral(value: string): string {
  if (!value) return value;
  if (
    value.length >= 2 &&
    value[0] === "'" &&
    (value[1] === "=" || value[1] === "+" || value[1] === "-" || value[1] === "@")
  ) {
    return value.slice(1);
  }
  return value;
}


/** Grid representation of a backup sheet (row-major string cells). */
export interface BackupSheetGrid {
  /** A1 magic */
  magic: string;
  headers: string[];
  rows: FormulaBackupRow[];
}

export function isBackupMagic(value: unknown): boolean {
  return String(value ?? "").trim() === FORMULA_BACKUP_MAGIC;
}

export function createBackupRows(
  cells: FormulaCellRecord[],
  options: {
    backupId: string;
    createdAt?: string;
    sourceRange?: string;
  },
): FormulaBackupRow[] {
  const createdAt = options.createdAt ?? new Date().toISOString();
  const sourceRange = options.sourceRange ?? "";
  const rows: FormulaBackupRow[] = [];
  for (const cell of cells) {
    if (!cell.formula.startsWith("=")) continue;
    rows.push({
      backupId: options.backupId,
      createdAt,
      sheet: cell.sheetName.trim(),
      address: normalizeA1Address(cell.address),
      formula: cell.formula,
      formulaR1C1: cell.formulaR1C1 ?? "",
      numberFormat: cell.numberFormat ?? "",
      locked: cell.locked === true,
      spillAddress: cell.spillAddress ?? "",
      sourceRange,
    });
  }
  return rows;
}

/** Serialize to a 2D values matrix: row1 magic, row2 headers, then data. */
export function encodeBackupSheet(rows: FormulaBackupRow[]): string[][] {
  const grid: string[][] = [
    [FORMULA_BACKUP_MAGIC],
    [...FORMULA_BACKUP_HEADERS],
  ];
  for (const row of rows) {
    grid.push([
      row.backupId,
      row.createdAt,
      row.sheet,
      row.address,
      encodeBackupLiteral(row.formula),
      encodeBackupLiteral(row.formulaR1C1),
      row.numberFormat,
      row.locked ? "1" : "0",
      row.spillAddress,
      row.sourceRange,
    ]);
  }
  return grid;
}

function cellAt(matrix: unknown[][], r: number, c: number): string {
  const row = matrix[r];
  if (!row) return "";
  const v = row[c];
  if (v == null) return "";
  return String(v);
}

function parseLocked(raw: string): boolean {
  const t = raw.trim().toLowerCase();
  return t === "1" || t === "true" || t === "yes";
}

/**
 * Decode a backup sheet matrix. Corrupted / incomplete data rows are skipped
 * (returned in `skipped`).
 */
export function decodeBackupSheet(matrix: unknown[][]): {
  ok: boolean;
  grid?: BackupSheetGrid;
  skipped: number[];
  error?: string;
} {
  if (!matrix.length) {
    return { ok: false, skipped: [], error: "empty backup sheet" };
  }
  const magic = cellAt(matrix, 0, 0).trim();
  if (!isBackupMagic(magic)) {
    return {
      ok: false,
      skipped: [],
      error: `invalid magic: ${magic || "(empty)"}`,
    };
  }

  const headerRow = matrix[1] ?? [];
  const headers = FORMULA_BACKUP_HEADERS.map((_, i) =>
    String(headerRow[i] ?? "").trim(),
  );
  const headerOk = FORMULA_BACKUP_HEADERS.every((h, i) => headers[i] === h);

  const rows: FormulaBackupRow[] = [];
  const skipped: number[] = [];
  // Data starts at row index 2 (Excel row 3)
  for (let r = 2; r < matrix.length; r += 1) {
    const backupId = cellAt(matrix, r, 0).trim();
    if (!backupId) {
      skipped.push(r);
      continue;
    }
    const sheet = cellAt(matrix, r, 2).trim();
    const address = normalizeA1Address(cellAt(matrix, r, 3));
    const formula = decodeBackupLiteral(cellAt(matrix, r, 4));
    if (!sheet || !address || !formula.startsWith("=")) {
      skipped.push(r);
      continue;
    }
    rows.push({
      backupId,
      createdAt: cellAt(matrix, r, 1),
      sheet,
      address,
      formula,
      formulaR1C1: decodeBackupLiteral(cellAt(matrix, r, 5)),
      numberFormat: cellAt(matrix, r, 6),
      locked: parseLocked(cellAt(matrix, r, 7)),
      spillAddress: cellAt(matrix, r, 8),
      sourceRange: cellAt(matrix, r, 9),
    });
  }

  return {
    ok: headerOk || rows.length > 0,
    grid: {
      magic: FORMULA_BACKUP_MAGIC,
      headers: headerOk ? [...FORMULA_BACKUP_HEADERS] : headers,
      rows,
    },
    skipped,
    error: headerOk ? undefined : "header mismatch (rows still parsed when possible)",
  };
}

export function summarizeBackups(rows: FormulaBackupRow[]): FormulaBackupSummary[] {
  const groups = new Map<string, FormulaBackupSummary>();
  for (const row of rows) {
    let summary = groups.get(row.backupId);
    if (!summary) {
      summary = {
        backupId: row.backupId,
        createdAt: row.createdAt,
        formulaCount: 0,
        sheets: [],
        sourceRanges: [],
      };
      groups.set(row.backupId, summary);
    }
    summary.formulaCount += 1;
    if (row.sheet && !summary.sheets.includes(row.sheet)) {
      summary.sheets.push(row.sheet);
    }
    if (row.sourceRange && !summary.sourceRanges.includes(row.sourceRange)) {
      summary.sourceRanges.push(row.sourceRange);
    }
    // Prefer earliest non-empty createdAt already set; if empty, take this row's
    if (!summary.createdAt && row.createdAt) summary.createdAt = row.createdAt;
  }
  return [...groups.values()].sort((a, b) =>
    b.createdAt.localeCompare(a.createdAt),
  );
}

export function planRestore(
  rows: FormulaBackupRow[],
  backupId: string,
): FormulaRestorePlan | { error: string } {
  const id = backupId.trim();
  if (!id) return { error: "backupId required" };
  const items: FormulaRestorePlanItem[] = [];
  for (const row of rows) {
    if (row.backupId !== id) continue;
    items.push({
      sheet: row.sheet,
      address: row.address,
      formula: row.formula,
      formulaR1C1: row.formulaR1C1,
      numberFormat: row.numberFormat,
      locked: row.locked,
      spillAddress: row.spillAddress,
      sourceRange: row.sourceRange,
    });
  }
  if (items.length === 0) return { error: "formula_backup_not_found" };
  return { backupId: id, items };
}
