/**
 * Collect formula-bearing cells for governance scopes (workbook/sheet/target).
 */
import { absoluteA1FromOrigin } from "./a1Address";
import {
  FORMULA_BACKUP_SHEET_PREFIX,
  MAX_GOVERNANCE_FORMULA_CELLS,
  type FormulaGovernanceScope,
} from "./formulaGovernanceTypes";
import type { FormulaCellRecord } from "../formulaGovernance";
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";

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

export async function collectFormulaCellsFromSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress: string | undefined,
  limitations: string[],
  sourceRangeHint?: string,
): Promise<FormulaCellRecord[]> {
  sheet.load("name");
  const range = await resolveGovernanceRange(context, sheet, rangeAddress);
  // formulas + values + numberFormat; formulaR1C1/locked optional via per-cell load later
  range.load("address,formulas,values,numberFormat,rowCount,columnCount");
  await context.sync();

  const sheetName = sheet.name;
  const origin = bareAddress(range.address || "A1");
  const formulas = Array.isArray(range.formulas) ? (range.formulas as unknown[][]) : [];
  const values = Array.isArray(range.values) ? (range.values as unknown[][]) : [];
  const nf = range.numberFormat;

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

  const records: FormulaCellRecord[] = [];
  const sourceRange =
    sourceRangeHint?.trim() ||
    (rangeAddress ? bareAddress(rangeAddress) : origin);

  // Batch-load locked + formulaR1C1 where available
  const cellProxies = capped.map((coord) => {
    const cell = range.getCell(coord.row, coord.col);
    cell.load("address,formulas,values,numberFormat");
    try {
      if (cell.format?.protection) cell.format.protection.load("locked");
    } catch {
      // ExcelApi 1.2 may be missing in some fakes — locked stays undefined.
    }
    return { coord, cell };
  });
  await context.sync();

  for (const { coord, cell } of cellProxies) {
    const formulaRaw = formulas[coord.row]?.[coord.col];
    const formula = isFormulaText(formulaRaw) ? formulaRaw : String(formulaRaw ?? "");
    if (!isFormulaText(formula)) continue;
    const address = absoluteA1FromOrigin(origin, coord.row, coord.col);
    let locked: boolean | undefined;
    try {
      locked = cell.format?.protection?.locked === true;
    } catch {
      locked = undefined;
    }
    records.push({
      sheetName,
      address,
      formula,
      value: values[coord.row]?.[coord.col],
      formulaR1C1: "",
      numberFormat: numberFormatAt(nf as string[][] | string, coord.row, coord.col),
      locked,
      spillAddress: "",
      // host-only annotation via sourceRange on backup rows
    });
  }

  // Attach sourceRange for backup helpers via a side channel is not on FormulaCellRecord;
  // callers pass sourceRange separately. Silence unused.
  void sourceRange;
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
      const part = await collectFormulaCellsFromSheet(context, ws, undefined, limitations);
      cells.push(...part);
    }
  } else {
    const sheet = context.workbook.worksheets.getItem(input.sheetName!);
    if (!sourceRange) sourceRange = input.scope === "sheet" ? "sheet" : input.range!;
    const part = await collectFormulaCellsFromSheet(
      context,
      sheet,
      input.scope === "target" ? input.range : undefined,
      limitations,
      sourceRange,
    );
    cells.push(...part);
  }

  return { cells, sourceRange };
}
