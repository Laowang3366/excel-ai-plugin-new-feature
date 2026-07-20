/**
 * Explicit mapping-based formula repair (desktop repairFormulaReferences).
 * No smart guessing — only apply provided find/replace pairs.
 */

import { makeCellId, normalizeA1Address } from "./address";
import type {
  FormulaCellRecord,
  FormulaRepairItem,
  FormulaRepairPlan,
  FormulaReplacement,
} from "./types";

export interface RepairOptions {
  /**
   * When false (default), only formulas containing #REF! are rewritten.
   * When true, apply mappings to every formula cell (desktop applyAllMappings).
   */
  applyAllMappings?: boolean;
}

function applyReplacements(
  formula: string,
  replacements: FormulaReplacement[],
): string {
  let after = formula;
  for (const item of replacements) {
    if (!item.find) continue;
    // desktop: string.Replace with Ordinal — non-regex, all occurrences
    after = after.split(item.find).join(item.replace);
  }
  return after;
}

function stillBroken(formula: string): boolean {
  return formula.toUpperCase().includes("#REF!");
}

/**
 * Produce a repair plan from in-memory cells + explicit replacements.
 * Hosts apply `repairs[].after` formulas; if `complete` is false, desktop
 * would refuse to save (formula_repair_incomplete).
 */
export function planFormulaRepairs(
  cells: FormulaCellRecord[],
  replacements: FormulaReplacement[],
  options: RepairOptions = {},
): FormulaRepairPlan {
  const applyAll = options.applyAllMappings === true;
  const repairs: FormulaRepairItem[] = [];
  const unresolved: FormulaRepairItem[] = [];

  for (const cell of cells) {
    const before = cell.formula;
    if (!before.startsWith("=")) continue;
    if (!applyAll && !stillBroken(before)) continue;

    const after = applyReplacements(before, replacements);
    if (after === before) {
      if (stillBroken(before)) {
        const item: FormulaRepairItem = {
          cell: makeCellId(cell.sheetName, normalizeA1Address(cell.address)),
          before,
          after,
          strategy: "mapping",
        };
        unresolved.push(item);
      }
      continue;
    }

    const item: FormulaRepairItem = {
      cell: makeCellId(cell.sheetName, normalizeA1Address(cell.address)),
      before,
      after,
      strategy: "mapping",
    };
    if (stillBroken(after)) unresolved.push(item);
    else repairs.push(item);
  }

  return {
    repairs,
    unresolved,
    repairedCount: repairs.length,
    unresolvedCount: unresolved.length,
    complete: unresolved.length === 0,
  };
}

/** Validate that a post-write formula set no longer contains #REF!. */
export function validateRepairedFormulas(
  cells: Array<{ cell: string; formula: string }>,
): { ok: true } | { ok: false; stillBroken: string[] } {
  const stillBroken = cells
    .filter((c) => stillBrokenFormula(c.formula))
    .map((c) => c.cell);
  if (stillBroken.length === 0) return { ok: true };
  return { ok: false, stillBroken };
}

function stillBrokenFormula(formula: string): boolean {
  return formula.toUpperCase().includes("#REF!");
}
