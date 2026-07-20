/**
 * Collect formula-bearing cells for governance scopes (workbook/sheet/target).
 * Backup metadata (locked/spill/formulaR1C1) is batched — never per-cell sync.
 */
import { absoluteA1FromOrigin } from "./a1Address";
import {
  FORMULA_BACKUP_SHEET_PREFIX,
  MAX_GOVERNANCE_FORMULA_CELLS,
  type FormulaGovernanceScope,
} from "./formulaGovernanceTypes";
import type { FormulaCellRecord } from "../formulaGovernance";
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";

export type CollectFormulaOptions = {
  /** locked / spill / formulaR1C1 for workbook backup (repair/convert). Default false. */
  includeBackupMetadata?: boolean;
};

export function requireGovernanceScope(
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

export function bareAddress(address: string): string {
  const bare = address.includes("!")
    ? address.slice(address.lastIndexOf("!") + 1)
    : address;
  return bare.replace(/\$/g, "").trim();
}

export function isBackupSheetName(name: string): boolean {
  return name.startsWith(FORMULA_BACKUP_SHEET_PREFIX);
}

export async function resolveGovernanceRange(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress?: string,
): Promise<ExcelRange> {
  if (rangeAddress && rangeAddress.trim()) {
    return sheet.getRange(bareAddress(rangeAddress));
  }
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("isNullObject,address");
  await context.sync();
  if (used.isNullObject) {
    return sheet.getRange("A1");
  }
  const address = typeof used.address === "string" ? used.address : "A1";
  return sheet.getRange(bareAddress(address) || "A1");
}

function isFormulaText(text: unknown): text is string {
  return typeof text === "string" && text.length > 1 && text.startsWith("=");
}

function numberFormatAt(nf: string[][] | string | undefined, r: number, c: number): string {
  if (typeof nf === "string") return nf;
  if (!Array.isArray(nf)) return "";
  const row = nf[r];
  if (!row) return "";
  const v = row[c];
  return v == null ? "" : String(v);
}

function isExcelApiSupported(version: string): boolean {
  if (typeof window === "undefined") return false;
  const isSetSupported = window.Office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return true;
  try {
    return Boolean(isSetSupported.call(window.Office?.context?.requirements, "ExcelApi", version));
  } catch {
    return false;
  }
}

function tryLoadFormulasR1C1(range: ExcelRange): boolean {
  const r = range as ExcelRange & { formulasR1C1?: unknown; load: (p: string) => void };
  try {
    r.load("formulasR1C1");
    return true;
  } catch {
    return false;
  }
}

type SpillProxy = ExcelRange & { isNullObject?: boolean };

export async function collectFormulaCellsFromSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress: string | undefined,
  limitations: string[],
  options: CollectFormulaOptions = {},
): Promise<FormulaCellRecord[]> {
  const includeMeta = options.includeBackupMetadata === true;
  sheet.load("name");
  const range = await resolveGovernanceRange(context, sheet, rangeAddress);
  const loadR1C1 = includeMeta ? tryLoadFormulasR1C1(range) : false;
  range.load(
    includeMeta
      ? "address,formulas,values,numberFormat,rowCount,columnCount"
      : "address,formulas,values,rowCount,columnCount",
  );
  await context.sync();

  if (includeMeta && !loadR1C1 && !limitations.some((l) => l.includes("formulaR1C1"))) {
    limitations.push("formulaR1C1 unavailable on Range (formulasR1C1 not loaded); stored empty");
  }

  const sheetName = sheet.name;
  const origin = bareAddress(range.address || "A1");
  const formulas = Array.isArray(range.formulas) ? (range.formulas as unknown[][]) : [];
  const values = Array.isArray(range.values) ? (range.values as unknown[][]) : [];
  const nf = includeMeta ? range.numberFormat : undefined;
  const r1c1Matrix = loadR1C1
    ? ((range as ExcelRange & { formulasR1C1?: unknown[][] }).formulasR1C1 ?? [])
    : [];

  const coords: { row: number; col: number }[] = [];
  for (let r = 0; r < formulas.length; r += 1) {
    const row = formulas[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      if (isFormulaText(row[c])) coords.push({ row: r, col: c });
    }
  }

  let capped = coords;
  if (coords.length > MAX_GOVERNANCE_FORMULA_CELLS) {
    limitations.push(
      `formula cell cap ${MAX_GOVERNANCE_FORMULA_CELLS}: collected first ${MAX_GOVERNANCE_FORMULA_CELLS} of ${coords.length}`,
    );
    capped = coords.slice(0, MAX_GOVERNANCE_FORMULA_CELLS);
  }

  if (!includeMeta) {
    return capped.map((coord) => {
      const formulaRaw = formulas[coord.row]?.[coord.col];
      const formula = isFormulaText(formulaRaw) ? formulaRaw : String(formulaRaw ?? "");
      return {
        sheetName,
        address: absoluteA1FromOrigin(origin, coord.row, coord.col),
        formula,
        value: values[coord.row]?.[coord.col],
        formulaR1C1: "",
        numberFormat: "",
        locked: undefined,
        spillAddress: "",
      };
    });
  }

  // Backup metadata path: queue all cell loads, then a single sync.
  type MetaPending = {
    coord: { row: number; col: number };
    cell: ExcelRange;
    spill: SpillProxy | null;
  };
  const pending: MetaPending[] = [];
  // Only ExcelApi 1.12+ null-object spill API is safe to batch. Never fall back to
  // getSpillingToRange for ordinary formulas — its errors often surface only at sync
  // and would fail the entire repair/convert batch.
  const canProbeSpill =
    isExcelApiSupported("1.12") &&
    typeof (range as ExcelRange & { getSpillingToRangeOrNullObject?: () => ExcelRange })
      .getSpillingToRangeOrNullObject === "function";

  if (!canProbeSpill && !limitations.some((l) => l.includes("spillAddress"))) {
    limitations.push(
      "spillAddress unavailable (requires ExcelApi 1.12 getSpillingToRangeOrNullObject); stored empty",
    );
  } else if (canProbeSpill && !limitations.some((l) => l.includes("spillAddress"))) {
    limitations.push("spillAddress probed via getSpillingToRangeOrNullObject (ExcelApi 1.12+)");
  }

  for (const coord of capped) {
    const cell = range.getCell(coord.row, coord.col);
    cell.load("address");
    try {
      if (cell.format?.protection) cell.format.protection.load("locked");
    } catch {
      // optional locked
    }
    let spill: SpillProxy | null = null;
    if (canProbeSpill) {
      const ext = cell as ExcelRange & { getSpillingToRangeOrNullObject?: () => ExcelRange };
      spill = ext.getSpillingToRangeOrNullObject!() as SpillProxy;
      spill.load("address,isNullObject");
    }
    pending.push({ coord, cell, spill });
  }

  if (pending.length > 0) {
    await context.sync();
  }

  const records: FormulaCellRecord[] = [];
  for (const item of pending) {
    const formulaRaw = formulas[item.coord.row]?.[item.coord.col];
    const formula = isFormulaText(formulaRaw) ? formulaRaw : String(formulaRaw ?? "");
    if (!isFormulaText(formula)) continue;

    let locked: boolean | undefined;
    try {
      locked = item.cell.format?.protection?.locked === true;
    } catch {
      locked = undefined;
    }

    let formulaR1C1 = "";
    const r1 = r1c1Matrix[item.coord.row]?.[item.coord.col];
    if (typeof r1 === "string" && r1.startsWith("=")) formulaR1C1 = r1;

    let spillAddress = "";
    if (item.spill) {
      const nullObj = item.spill.isNullObject === true;
      if (!nullObj && typeof item.spill.address === "string" && item.spill.address) {
        spillAddress = bareAddress(item.spill.address);
      }
    }

    records.push({
      sheetName,
      address: absoluteA1FromOrigin(origin, item.coord.row, item.coord.col),
      formula,
      value: values[item.coord.row]?.[item.coord.col],
      formulaR1C1,
      numberFormat: numberFormatAt(nf as string[][] | string, item.coord.row, item.coord.col),
      locked,
      spillAddress,
    });
  }

  return records;
}

export async function collectFormulaCells(
  context: ExcelRequestContext,
  input: {
    scope: FormulaGovernanceScope;
    sheetName?: string;
    range?: string;
  },
  limitations: string[],
  options: CollectFormulaOptions = {},
): Promise<{ cells: FormulaCellRecord[]; sourceRange: string }> {
  requireGovernanceScope(input.scope, input.sheetName, input.range);
  const cells: FormulaCellRecord[] = [];
  let sourceRange = input.range?.trim() || "";

  if (input.scope === "workbook") {
    context.workbook.worksheets.load("items/name");
    await context.sync();
    sourceRange = "workbook";
    for (const ws of context.workbook.worksheets.items) {
      ws.load("name");
      await context.sync();
      if (isBackupSheetName(ws.name)) continue;
      cells.push(
        ...(await collectFormulaCellsFromSheet(context, ws, undefined, limitations, options)),
      );
    }
  } else {
    const sheet = context.workbook.worksheets.getItem(input.sheetName!);
    if (!sourceRange) sourceRange = input.scope === "sheet" ? "sheet" : input.range!;
    cells.push(
      ...(await collectFormulaCellsFromSheet(
        context,
        sheet,
        input.scope === "target" ? input.range : undefined,
        limitations,
        options,
      )),
    );
  }

  return { cells, sourceRange };
}
