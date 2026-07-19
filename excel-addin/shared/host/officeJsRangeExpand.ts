import { absoluteA1FromOrigin, isSingleCellAddress } from "./a1Address";
import type { ExcelRange, ExcelRequestContext } from "./officeJsRuntime";
import {
  normalizeFormulas,
  normalizeMatrix,
  withExcel,
} from "./officeJsRuntime";
import type {
  FormulaContextData,
  FormulaContextEntry,
  HostResult,
  RangeData,
  RangeExpandMode,
} from "./types";

/**
 * Desktop parity: omitted expand on single-cell → spill; explicit none stays none.
 */
function resolveExpandMode(
  address: string,
  expand: RangeExpandMode | undefined,
): RangeExpandMode {
  if (expand !== undefined) return expand;
  return isSingleCellAddress(address) ? "spill" : "none";
}

function resolveExpandedRange(
  source: ExcelRange,
  mode: RangeExpandMode,
): { range: ExcelRange; expandMode: RangeExpandMode; expanded: boolean } {
  if (mode === "spill") {
    return { range: source.getSpillingToRange(), expandMode: "spill", expanded: true };
  }
  if (mode === "currentRegion") {
    return {
      range: source.getSurroundingRegion(),
      expandMode: "currentRegion",
      expanded: true,
    };
  }
  if (mode === "currentArray") {
    return { range: source.getCurrentArray(), expandMode: "currentArray", expanded: true };
  }
  return { range: source, expandMode: "none", expanded: false };
}

export async function officeJsReadRange(
  sheetName: string,
  address: string,
  expand?: RangeExpandMode,
): Promise<HostResult<RangeData>> {
  return withExcel("range.read", async (context: ExcelRequestContext) => {
    const mode = resolveExpandMode(address, expand);
    const source = context.workbook.worksheets.getItem(sheetName).getRange(address);
    let resolved;
    try {
      resolved = resolveExpandedRange(source, mode);
    } catch {
      resolved = { range: source, expandMode: mode, expanded: false };
    }
    resolved.range.load("address,values,formulas");
    await context.sync();
    return {
      sheetName,
      address: resolved.range.address,
      values: normalizeMatrix(resolved.range.values),
      formulas: normalizeFormulas(resolved.range.formulas),
      expanded: resolved.expanded,
      expandMode: resolved.expandMode,
    };
  });
}

export async function officeJsGetFormulaContext(
  sheetName: string,
  address?: string,
): Promise<HostResult<FormulaContextData>> {
  return withExcel("formula.context", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const useUsedRange = !address || !address.trim();
    if (useUsedRange) {
      const used = sheet.getUsedRangeOrNullObject(true);
      // isNullObject is only valid on NullObject ranges (UsedRange path).
      used.load("address,values,formulas,rowCount,columnCount,isNullObject");
      await context.sync();
      if (used.isNullObject) {
        return { sheetName, address: "", formulas: [], cells: [] };
      }
      return collectFormulaEntries(context, sheetName, used);
    }
    const range = sheet.getRange(address!);
    // Ordinary Range: do not load isNullObject.
    range.load("address,values,formulas,rowCount,columnCount");
    await context.sync();
    return collectFormulaEntries(context, sheetName, range);
  });
}

/**
 * Batch-load all formula cell addresses, then ONE sync, then read A1 addresses.
 * Matches real Office.js: properties set by load() are only available after sync().
 */
async function collectFormulaEntries(
  context: ExcelRequestContext,
  sheetName: string,
  range: ExcelRange,
): Promise<FormulaContextData> {
  const formulasMatrix = normalizeFormulas(range.formulas);
  const values = normalizeMatrix(range.values);
  const originBare = range.address.includes("!")
    ? range.address.split("!")[1]!
    : range.address;

  type Pending = {
    row: number;
    col: number;
    formula: string;
    value: import("./types").CellValue;
    cell: ExcelRange;
  };
  const pending: Pending[] = [];

  for (let r = 0; r < formulasMatrix.length; r += 1) {
    for (let c = 0; c < (formulasMatrix[r]?.length ?? 0); c += 1) {
      const formula = formulasMatrix[r][c] ?? "";
      if (!formula.startsWith("=")) continue;
      const cell = range.getCell(r, c);
      cell.load("address");
      pending.push({
        row: r,
        col: c,
        formula,
        value: values[r]?.[c] ?? null,
        cell,
      });
    }
  }

  if (pending.length > 0) {
    await context.sync();
  }

  const formulas: FormulaContextEntry[] = pending.map((item) => {
    const hostAddr = String(item.cell.address ?? "").replace(/^.*!/, "");
    // Require real A1 after sync; fall back to absolute A1 from origin only if host empty.
    const a1 =
      hostAddr && /^[A-Z]+\d+$/i.test(hostAddr)
        ? hostAddr.toUpperCase()
        : absoluteA1FromOrigin(originBare, item.row, item.col);
    return {
      address: a1,
      formula: item.formula,
      value: item.value,
    };
  });

  return {
    sheetName,
    address: range.address.replace(/^.*!/, ""),
    formulas,
    cells: formulas,
  };
}
