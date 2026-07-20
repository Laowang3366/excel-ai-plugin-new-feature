/**
 * WPS JSA formula governance helpers (collect + backup sheet).
 */
import {
  createBackupRows,
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
import {
  formulaMatrixFrom,
  getSheet,
  matrixFrom,
  type WpsRange,
  type WpsSheet,
  type WpsWorkbook,
} from "./wpsJsaRuntime";

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

export function listSheets(workbook: WpsWorkbook): WpsSheet[] {
  const sheets = workbook.Worksheets;
  if (!sheets || typeof sheets.Count !== "number" || typeof sheets.Item !== "function") {
    throw new Error("Worksheets collection unavailable");
  }
  const out: WpsSheet[] = [];
  for (let i = 1; i <= sheets.Count; i += 1) out.push(sheets.Item(i));
  return out;
}

export function isBackupName(name: string): boolean {
  return name.startsWith(FORMULA_BACKUP_SHEET_PREFIX);
}

export function resolveRange(sheet: WpsSheet, rangeAddress?: string): WpsRange {
  if (!sheet.Range) throw new Error("Sheet.Range unavailable");
  if (rangeAddress?.trim()) return sheet.Range(bare(rangeAddress));
  if (!sheet.UsedRange?.Address) return sheet.Range("A1");
  return sheet.Range(bare(String(sheet.UsedRange.Address)));
}

export function collectFromSheet(
  sheet: WpsSheet,
  rangeAddress: string | undefined,
  limitations: string[],
): FormulaCellRecord[] {
  const range = resolveRange(sheet, rangeAddress);
  const formulas = formulaMatrixFrom(range.Formula);
  const values = matrixFrom(range.Value2);
  const origin = bare(String(range.Address ?? rangeAddress ?? "A1"));
  const records: FormulaCellRecord[] = [];
  for (let r = 0; r < formulas.length; r += 1) {
    for (let c = 0; c < (formulas[r]?.length ?? 0); c += 1) {
      const formula = formulas[r]![c] ?? "";
      if (!formula.startsWith("=")) continue;
      let numberFormat = "";
      let locked: boolean | undefined;
      try {
        const cell = sheet.Range(absoluteA1FromOrigin(origin, r, c));
        if (cell.NumberFormat != null) numberFormat = String(cell.NumberFormat);
        const lockedProp = (cell as WpsRange & { Locked?: boolean }).Locked;
        if (typeof lockedProp === "boolean") locked = lockedProp;
      } catch {
        // optional metadata
      }
      records.push({
        sheetName: sheet.Name,
        address: absoluteA1FromOrigin(origin, r, c),
        formula,
        value: values[r]?.[c],
        formulaR1C1: "",
        numberFormat,
        locked,
        spillAddress: "",
      });
    }
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

export function findBackupSheet(
  workbook: WpsWorkbook,
  create: boolean,
): { sheet: WpsSheet | null; limitations: string[] } {
  const limitations: string[] = [];
  for (const sheet of listSheets(workbook)) {
    if (!isBackupName(sheet.Name)) continue;
    try {
      const a1 = sheet.Range("A1");
      const magic = String(matrixFrom(a1.Value2)?.[0]?.[0] ?? "").trim();
      if (magic === FORMULA_BACKUP_MAGIC) return { sheet, limitations };
      limitations.push(`sheet ${sheet.Name} has backup prefix but invalid magic; left untouched`);
    } catch {
      limitations.push(`cannot read magic on ${sheet.Name}`);
    }
  }
  if (!create) return { sheet: null, limitations };
  const sheets = workbook.Worksheets;
  if (typeof sheets.Add !== "function") {
    throw new Error("Worksheets.Add unavailable for backup sheet");
  }
  const stamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const name = `${FORMULA_BACKUP_SHEET_PREFIX}${stamp}`.slice(0, 31);
  const sheet = sheets.Add();
  sheet.Name = name;
  sheet.Range("A1").Value2 = [[FORMULA_BACKUP_MAGIC]];
  sheet.Range("A2:J2").Value2 = [[...FORMULA_BACKUP_HEADERS]];
  const vis = sheet as WpsSheet & { Visible?: number | string };
  if ("Visible" in vis) vis.Visible = 2;
  limitations.push(`created backup sheet ${name}`);
  return { sheet, limitations };
}

export function readBackupMatrix(sheet: WpsSheet): unknown[][] {
  if (!sheet.UsedRange) {
    return [[FORMULA_BACKUP_MAGIC], [...FORMULA_BACKUP_HEADERS]];
  }
  return matrixFrom(sheet.UsedRange.Value2) as unknown[][];
}

export function appendBackup(
  sheet: WpsSheet,
  cells: FormulaCellRecord[],
  backupId: string,
  sourceRange: string,
): void {
  const rows = createBackupRows(cells, { backupId, sourceRange });
  if (rows.length === 0) return;
  let start = 3;
  try {
    const used = sheet.UsedRange;
    if (used?.Address) {
      const bareAddr = bare(String(used.Address));
      const m = /:.*?(\d+)$/.exec(bareAddr) || /^[A-Z]+(\d+)$/i.exec(bareAddr);
      if (m) start = Math.max(3, Number(m[1]) + 1);
    }
  } catch {
    start = 3;
  }
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
  const end = start + grid.length - 1;
  sheet.Range(`A${start}:J${end}`).Value2 = grid;
  const vis = sheet as WpsSheet & { Visible?: number | string };
  if ("Visible" in vis) vis.Visible = 2;
}
