/**
 * WPS JSA formula governance helpers (collect + backup sheet).
 * Formula cells written as text: NumberFormat "@" + apostrophe encode.
 * Backup sheet must be VeryHidden (Visible=2) or operations fail closed.
 */
import {
  createBackupRows,
  encodeBackupLiteral,
  encodeBackupSheet,
  FORMULA_BACKUP_HEADERS,
  FORMULA_BACKUP_MAGIC,
  type FormulaCellRecord,
} from "../formulaGovernance";
import { absoluteA1FromOrigin } from "./a1Address";
import {
  FORMULA_BACKUP_SHEET_PREFIX,
  MAX_GOVERNANCE_FORMULA_CELLS,
  type FormulaGovernanceScope,
} from "./formulaGovernanceTypes";
import { readWpsAddress } from "./wpsJsaAddress";
import {
  formulaMatrixFrom,
  getSheet,
  matrixFrom,
  type WpsRange,
  type WpsSheet,
  type WpsWorkbook,
} from "./wpsJsaRuntime";

export type WpsSheetExt = WpsSheet & {
  Visible?: number | string;
  Name: string;
};

export type WpsRangeExt = WpsRange & {
  FormulaR1C1?: unknown;
  Locked?: boolean;
  NumberFormat?: string | unknown;
};

export function bare(address: string): string {
  const a = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  return a.replace(/\$/g, "").trim();
}

export function requireScope(
  scope: FormulaGovernanceScope,
  sheetName?: string,
  range?: string,
): void {
  if (scope !== "workbook" && scope !== "sheet" && scope !== "target") {
    throw new Error("scope must be workbook|sheet|target");
  }
  if (scope !== "workbook" && (!sheetName || !sheetName.trim())) {
    throw new Error("sheetName is required for scope sheet|target");
  }
  if (scope === "target" && (!range || !range.trim())) {
    throw new Error("range is required for scope=target");
  }
}

export function newId(): string {
  const c = globalThis.crypto as Crypto | undefined;
  if (c?.randomUUID) return c.randomUUID().replace(/-/g, "");
  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2, 10)}`;
}

export function listSheets(workbook: WpsWorkbook): WpsSheetExt[] {
  const sheets = workbook.Worksheets;
  if (!sheets || typeof sheets.Count !== "number" || typeof sheets.Item !== "function") {
    throw new Error("Worksheets collection unavailable");
  }
  const out: WpsSheetExt[] = [];
  for (let i = 1; i <= sheets.Count; i += 1) out.push(sheets.Item(i) as WpsSheetExt);
  return out;
}

export function isBackupName(name: string): boolean {
  return name.startsWith(FORMULA_BACKUP_SHEET_PREFIX);
}

export function resolveRange(sheet: WpsSheet, rangeAddress?: string): WpsRangeExt {
  if (!sheet.Range) throw new Error("Sheet.Range unavailable");
  if (rangeAddress?.trim()) return sheet.Range(bare(rangeAddress)) as WpsRangeExt;
  const usedAddr = readWpsAddress(sheet.UsedRange);
  if (!usedAddr) return sheet.Range("A1") as WpsRangeExt;
  return sheet.Range(bare(usedAddr)) as WpsRangeExt;
}


export function writeTextMatrix(range: WpsRangeExt, matrix: unknown[][]): void {
  try {
    range.NumberFormat = "@";
  } catch {
    // optional
  }
  const encoded = matrix.map((row) =>
    row.map((cell) => (typeof cell === "string" ? encodeBackupLiteral(cell) : cell)),
  );
  // Only encode formula-like on individual cells that need it; for headers/ids don't double-encode
  // encodeBackupLiteral is no-op for non-formula-like — good for headers.
  range.Value2 = encoded;
}

export function ensureVeryHidden(sheet: WpsSheetExt): void {
  if (!("Visible" in sheet) || sheet.Visible === undefined) {
    throw new Error("backup_sheet_visibility_unavailable: Visible member missing");
  }
  sheet.Visible = 2;
  const v = sheet.Visible;
  const asText = String(v ?? "").toLowerCase();
  const ok =
    v === 2 ||
    asText === "2" ||
    asText === "veryhidden" ||
    asText === "xlsheetveryhidden";
  if (!ok) {
    throw new Error(`backup_sheet_visibility_unavailable: expected VeryHidden, got ${String(v)}`);
  }
}


export function canRemoveBackupRows(sheet: WpsSheetExt): boolean {
  const used = sheet.UsedRange as (WpsRange & { Clear?: () => void }) | undefined;
  return Boolean(used && typeof used.Clear === "function");
}

export function collectFromSheet(
  sheet: WpsSheet,
  rangeAddress: string | undefined,
  limitations: string[],
): FormulaCellRecord[] {
  const range = resolveRange(sheet, rangeAddress);
  const formulas = formulaMatrixFrom(range.Formula);
  const values = matrixFrom(range.Value2);
  const origin = bare(readWpsAddress(range, rangeAddress ?? "A1") ?? rangeAddress ?? "A1");
  let r1c1Available = false;
  const records: FormulaCellRecord[] = [];
  for (let r = 0; r < formulas.length; r += 1) {
    for (let c = 0; c < (formulas[r]?.length ?? 0); c += 1) {
      const formula = formulas[r]![c] ?? "";
      if (!formula.startsWith("=")) continue;
      let numberFormat = "";
      let locked: boolean | undefined;
      let formulaR1C1 = "";
      try {
        const cell = sheet.Range(absoluteA1FromOrigin(origin, r, c)) as WpsRangeExt;
        if (cell.NumberFormat != null) numberFormat = String(cell.NumberFormat);
        if (typeof cell.Locked === "boolean") locked = cell.Locked;
        if (cell.FormulaR1C1 != null) {
          const r1 = String(formulaMatrixFrom(cell.FormulaR1C1)?.[0]?.[0] ?? cell.FormulaR1C1);
          if (r1.startsWith("=")) {
            formulaR1C1 = r1;
            r1c1Available = true;
          }
        }
      } catch {
        // optional metadata
      }
      records.push({
        sheetName: sheet.Name,
        address: absoluteA1FromOrigin(origin, r, c),
        formula,
        value: values[r]?.[c],
        formulaR1C1,
        numberFormat,
        locked,
        spillAddress: "",
      });
    }
  }
  if (!r1c1Available && !limitations.some((l) => l.includes("formulaR1C1"))) {
    limitations.push("formulaR1C1 unavailable on WPS Range; stored empty when member missing");
  }
  if (!limitations.some((l) => l.includes("spillAddress"))) {
    limitations.push("spillAddress not available on WPS JSA path; stored empty");
  }
  if (records.length > MAX_GOVERNANCE_FORMULA_CELLS) {
    limitations.push(
      `formula cell cap ${MAX_GOVERNANCE_FORMULA_CELLS}: collected first ${MAX_GOVERNANCE_FORMULA_CELLS}`,
    );
    return records.slice(0, MAX_GOVERNANCE_FORMULA_CELLS);
  }
  return records;
}

export function collectAll(
  workbook: WpsWorkbook,
  input: { scope: FormulaGovernanceScope; sheetName?: string; range?: string },
  limitations: string[],
): { cells: FormulaCellRecord[]; sourceRange: string } {
  requireScope(input.scope, input.sheetName, input.range);
  if (input.scope === "workbook") {
    const cells: FormulaCellRecord[] = [];
    for (const sheet of listSheets(workbook)) {
      if (isBackupName(sheet.Name)) continue;
      cells.push(...collectFromSheet(sheet, undefined, limitations));
    }
    return { cells, sourceRange: "workbook" };
  }
  const sheet = getSheet(workbook, input.sheetName!);
  if (!sheet) throw new Error(`Sheet "${input.sheetName}" not found`);
  const range = input.scope === "target" ? input.range : undefined;
  return {
    cells: collectFromSheet(sheet, range, limitations),
    sourceRange: input.range?.trim() || input.scope,
  };
}

export function probeSheetVisibilitySupport(workbook: WpsWorkbook): void {
  const sheets = listSheets(workbook);
  for (const sheet of sheets) {
    if ("Visible" in sheet && sheet.Visible !== undefined) return;
  }
  if (sheets.length > 0) {
    throw new Error("backup_sheet_visibility_unavailable: Visible member missing on worksheets");
  }
}

export function tryDeleteSheet(workbook: WpsWorkbook, sheet: WpsSheetExt): boolean {
  const nameBefore = String(sheet.Name ?? "");
  const del = (sheet as WpsSheetExt & { Delete?: () => void }).Delete;
  if (typeof del === "function") {
    try {
      del.call(sheet);
    } catch {
      // fall through to Remove / verify
    }
  }
  // Best-effort: drop from collection if host exposes custom remove (tests).
  const sheetsApi = workbook.Worksheets as WpsWorkbook["Worksheets"] & {
    Remove?: (s: WpsSheetExt) => void;
  };
  if (typeof sheetsApi.Remove === "function") {
    try {
      sheetsApi.Remove(sheet);
    } catch {
      // verify below
    }
  }

  // Must re-enumerate — Delete/Remove no-op must not count as cleaned up.
  try {
    const remaining = listSheets(workbook);
    const stillPresent = remaining.some(
      (s) => s === sheet || (nameBefore.length > 0 && s.Name === nameBefore),
    );
    return !stillPresent;
  } catch {
    return false;
  }
}

export function findBackupSheet(
  workbook: WpsWorkbook,
  create: boolean,
): { sheet: WpsSheetExt | null; limitations: string[] } {
  const limitations: string[] = [];
  for (const sheet of listSheets(workbook)) {
    if (!isBackupName(sheet.Name)) continue;
    try {
      const a1 = sheet.Range("A1") as WpsRangeExt;
      const raw = String(matrixFrom(a1.Value2)?.[0]?.[0] ?? a1.Value2 ?? "").trim();
      // magic is not formula-like after decode strip
      const magic = raw.startsWith("'") ? raw.slice(1) : raw;
      if (magic === FORMULA_BACKUP_MAGIC) return { sheet, limitations };
      limitations.push(`sheet ${sheet.Name} has backup prefix but invalid magic; left untouched`);
    } catch {
      limitations.push(`cannot read magic on ${sheet.Name}`);
    }
  }
  if (!create) return { sheet: null, limitations };

  // Precheck before Add so we never leave a visible orphan when Visible is missing.
  probeSheetVisibilitySupport(workbook);

  const sheets = workbook.Worksheets;
  if (typeof sheets.Add !== "function") {
    throw new Error("Worksheets.Add unavailable for backup sheet");
  }
  const stamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const name = `${FORMULA_BACKUP_SHEET_PREFIX}${stamp}`.slice(0, 31);
  const sheet = sheets.Add() as WpsSheetExt;
  try {
    // Hide BEFORE writing protocol content
    ensureVeryHidden(sheet);
    sheet.Name = name;
    // Single rectangular write A1:J2 so UsedRange captures magic+headers together.
    const boot = [
      [FORMULA_BACKUP_MAGIC, "", "", "", "", "", "", "", "", ""],
      [...FORMULA_BACKUP_HEADERS],
    ];
    writeTextMatrix(sheet.Range("A1:J2") as WpsRangeExt, boot);
    ensureVeryHidden(sheet);
    limitations.push(`created backup sheet ${name}`);
    return { sheet, limitations };
  } catch (error) {
    const deleted = tryDeleteSheet(workbook, sheet);
    const msg = error instanceof Error ? error.message : String(error);
    if (!deleted) {
      throw new Error(
        `backup_sheet_cleanup_failed: could not delete orphan sheet after: ${msg}`,
      );
    }
    throw error instanceof Error ? error : new Error(msg);
  }
}

export function readBackupMatrix(sheet: WpsSheet): unknown[][] {
  if (!sheet.UsedRange) {
    return [[FORMULA_BACKUP_MAGIC], [...FORMULA_BACKUP_HEADERS]];
  }
  return matrixFrom(sheet.UsedRange.Value2) as unknown[][];
}

export function appendBackup(
  sheet: WpsSheetExt,
  cells: FormulaCellRecord[],
  backupId: string,
  sourceRange: string,
): void {
  ensureVeryHidden(sheet);
  const rows = createBackupRows(cells, { backupId, sourceRange });
  if (rows.length === 0) return;
  let start = 3;
  try {
    const used = sheet.UsedRange;
    const usedAddr = readWpsAddress(used);
    if (usedAddr) {
      const bareAddr = bare(usedAddr);
      const m = /:.*?(\d+)$/.exec(bareAddr) || /^[A-Z]+(\d+)$/i.exec(bareAddr);
      if (m) start = Math.max(3, Number(m[1]) + 1);
    }
  } catch {
    start = 3;
  }
  const full = encodeBackupSheet(rows);
  const data = full.slice(2);
  const end = start + data.length - 1;
  writeTextMatrix(sheet.Range(`A${start}:J${end}`) as WpsRangeExt, data);
  ensureVeryHidden(sheet);
}

export function rewriteBackupSheet(sheet: WpsSheetExt, rows: import("../formulaGovernance").FormulaBackupRow[]): void {
  if (!canRemoveBackupRows(sheet)) {
    throw new Error("backup_row_delete_unavailable");
  }
  const used = sheet.UsedRange as WpsRange & { Clear?: () => void };
  used.Clear!();
  const grid = encodeBackupSheet(rows);
  const cols = 10;
  const padded = grid.map((row) => {
    const out = row.map((v) => (v == null ? "" : String(v)));
    while (out.length < cols) out.push("");
    return out.slice(0, cols);
  });
  writeTextMatrix(sheet.Range(`A1:J${Math.max(padded.length, 2)}`) as WpsRangeExt, padded);
  ensureVeryHidden(sheet);
}
