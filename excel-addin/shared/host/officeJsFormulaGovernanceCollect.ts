/**
 * Collect formula-bearing cells for governance scopes (workbook/sheet/target).
 * Probes formulasR1C1 / FormulaR1C1 and spill address when available.
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

function tryLoadFormulasR1C1(range: ExcelRange): boolean {
  const r = range as ExcelRange & { formulasR1C1?: unknown; load: (p: string) => void };
  try {
    r.load("formulasR1C1");
    return true;
  } catch {
    return false;
  }
}

export async function collectFormulaCellsFromSheet(
  context: ExcelRequestContext,
  sheet: ExcelWorksheet,
  rangeAddress: string | undefined,
  limitations: string[],
): Promise<FormulaCellRecord[]> {
  sheet.load("name");
  const range = await resolveGovernanceRange(context, sheet, rangeAddress);
  const loadR1C1 = tryLoadFormulasR1C1(range);
  range.load("address,formulas,values,numberFormat,rowCount,columnCount");
  await context.sync();

  if (!loadR1C1 && !limitations.some((l) => l.includes("formulaR1C1"))) {
    limitations.push("formulaR1C1 unavailable on Range (formulasR1C1 not loaded); stored empty");
  }

  const sheetName = sheet.name;
  const origin = bareAddress(range.address || "A1");
  const formulas = Array.isArray(range.formulas) ? (range.formulas as unknown[][]) : [];
  const values = Array.isArray(range.values) ? (range.values as unknown[][]) : [];
  const nf = range.numberFormat;
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

  let spillProbed = false;
  let spillUnavailable = false;
  const records: FormulaCellRecord[] = [];

  for (const coord of capped) {
    const formulaRaw = formulas[coord.row]?.[coord.col];
    const formula = isFormulaText(formulaRaw) ? formulaRaw : String(formulaRaw ?? "");
    if (!isFormulaText(formula)) continue;
    const address = absoluteA1FromOrigin(origin, coord.row, coord.col);
    const cell = range.getCell(coord.row, coord.col);
    cell.load("address,formulas,values,numberFormat");
    try {
      if (cell.format?.protection) cell.format.protection.load("locked");
    } catch {
      // optional
    }
    await context.sync();

    let locked: boolean | undefined;
    try {
      locked = cell.format?.protection?.locked === true;
    } catch {
      locked = undefined;
    }

    let formulaR1C1 = "";
    const r1 = r1c1Matrix[coord.row]?.[coord.col];
    if (typeof r1 === "string" && r1.startsWith("=")) {
      formulaR1C1 = r1;
    }

    let spillAddress = "";
    if (!spillUnavailable) {
      try {
        const spilling = cell.getSpillingToRange?.();
        if (spilling) {
          spilling.load("address,isNullObject");
          await context.sync();
          spillProbed = true;
          const nullObj = (spilling as ExcelRange & { isNullObject?: boolean }).isNullObject;
          if (nullObj !== true && typeof spilling.address === "string") {
            spillAddress = bareAddress(spilling.address);
          }
        } else {
          spillUnavailable = true;
        }
      } catch {
        spillUnavailable = true;
      }
    }

    records.push({
      sheetName,
      address,
      formula,
      value: values[coord.row]?.[coord.col],
      formulaR1C1,
      numberFormat: numberFormatAt(nf as string[][] | string, coord.row, coord.col),
      locked,
      spillAddress,
    });
  }

  if (spillUnavailable && !limitations.some((l) => l.includes("spillAddress"))) {
    limitations.push(
      "spillAddress unavailable (getSpillingToRange missing or failed); stored empty",
    );
  } else if (spillProbed && !limitations.some((l) => l.includes("spillAddress probed"))) {
    limitations.push("spillAddress probed via getSpillingToRange when host supports it");
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
      cells.push(...(await collectFormulaCellsFromSheet(context, ws, undefined, limitations)));
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
      )),
    );
  }

  return { cells, sourceRange };
}
