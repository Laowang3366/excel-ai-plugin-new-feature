/**
 * Office.js formula governance: dependencies inspect, references repair, convertToValues.
 * Write/readback paths batch queue + chunked sync (O(chunks), not O(cells)).
 */
import {
  buildDependencyReport,
  planFormulaRepairs,
  validateRepairedFormulas,
  type FormulaCellRecord,
} from "../formulaGovernance";
import type {
  FormulaConvertToValuesInfo,
  FormulaConvertToValuesInput,
  FormulaDependenciesInspectInfo,
  FormulaDependenciesInspectInput,
  FormulaReferencesRepairInfo,
  FormulaReferencesRepairInput,
} from "./formulaGovernanceTypes";
import { newBackupId, writeFormulaBackup } from "./officeJsFormulaGovernanceBackup";
import {
  bareAddress,
  collectFormulaCells,
  requireGovernanceScope,
} from "./officeJsFormulaGovernanceCollect";
import type { ExcelRange, ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail } from "./types";

/** Max cells queued per context.sync for governance writes/readbacks. */
export const GOVERNANCE_WRITE_CHUNK = 250;

function isExcelApi12Supported(): boolean {
  if (typeof window === "undefined") return false;
  const isSetSupported = window.Office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return true;
  try {
    return Boolean(isSetSupported.call(window.Office?.context?.requirements, "ExcelApi", "1.2"));
  } catch {
    return false;
  }
}

function matchCell(cells: FormulaCellRecord[], cellId: string): FormulaCellRecord | undefined {
  const key = cellId.toLowerCase();
  return cells.find((c) => `${c.sheetName}!${c.address}`.toLowerCase() === key);
}

function parseCellId(cellId: string): { sheetName: string; address: string } {
  const bang = cellId.lastIndexOf("!");
  return { sheetName: cellId.slice(0, bang), address: cellId.slice(bang + 1) };
}

/** Queue work in fixed-size chunks; one sync per chunk (not per cell). */
export async function forEachChunkSync<T>(
  context: ExcelRequestContext,
  items: readonly T[],
  chunkSize: number,
  apply: (item: T, index: number) => void,
): Promise<void> {
  const size = Math.max(1, chunkSize);
  for (let i = 0; i < items.length; i += size) {
    const end = Math.min(i + size, items.length);
    for (let j = i; j < end; j += 1) apply(items[j]!, j);
    await context.sync();
  }
}

export async function officeJsInspectFormulaDependencies(
  input: FormulaDependenciesInspectInput,
): Promise<HostResult<FormulaDependenciesInspectInfo>> {
  if (!isExcelApi12Supported()) {
    return fail(
      "formula.dependencies.inspect",
      "office-js",
      "ExcelApi 1.2 is not supported in this host",
    );
  }
  requireGovernanceScope(input.scope, input.sheetName, input.range);
  return withExcel("formula.dependencies.inspect", async (context) => {
    const limitations: string[] = [];
    const { cells } = await collectFormulaCells(context, input, limitations, {
      includeBackupMetadata: false,
    });
    const report = buildDependencyReport(cells);
    return {
      scope: input.scope,
      report,
      limitations: [...limitations, ...report.limitations],
    };
  });
}

export async function officeJsRepairFormulaReferences(
  input: FormulaReferencesRepairInput,
): Promise<HostResult<FormulaReferencesRepairInfo>> {
  if (!isExcelApi12Supported()) {
    return fail(
      "formula.references.repair",
      "office-js",
      "ExcelApi 1.2 is not supported in this host",
    );
  }
  requireGovernanceScope(input.scope, input.sheetName, input.range);
  if (!Array.isArray(input.replacements) || input.replacements.length === 0) {
    return fail("formula.references.repair", "office-js", "replacements required");
  }

  return withExcel("formula.references.repair", async (context) => {
    const limitations: string[] = [];
    const { cells, sourceRange } = await collectFormulaCells(context, input, limitations, {
      includeBackupMetadata: true,
    });
    const plan = planFormulaRepairs(cells, input.replacements, {
      applyAllMappings: input.applyAllMappings === true,
    });

    if (!plan.complete) {
      throw Object.assign(
        new Error("formula_repair_incomplete: still contains #REF! after mapping; no cells written"),
        {
          detail: {
            code: "formula_repair_incomplete",
            repairs: plan.repairs,
            unresolved: plan.unresolved,
            repairedCount: plan.repairedCount,
            unresolvedCount: plan.unresolvedCount,
          },
        },
      );
    }

    if (plan.repairs.length === 0) {
      return {
        scope: input.scope,
        backupId: "",
        repairs: [],
        unresolved: [],
        repairedCount: 0,
        unresolvedCount: 0,
        verified: true,
        limitations: [...limitations, "no matching formula cells to repair"],
      };
    }

    const backupByPlan: FormulaCellRecord[] = [];
    for (const repair of plan.repairs) {
      const hit = matchCell(cells, repair.cell);
      if (hit) backupByPlan.push(hit);
    }
    const backupId = newBackupId();
    await writeFormulaBackup(context, backupByPlan, {
      backupId,
      sourceRange,
      create: true,
    });

    // Batch write formulas, then batch readback — O(chunks) syncs.
    await forEachChunkSync(context, plan.repairs, GOVERNANCE_WRITE_CHUNK, (repair) => {
      const { sheetName, address } = parseCellId(repair.cell);
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(bareAddress(address));
      range.formulas = [[repair.after]];
    });

    const readTargets: Array<{ cell: string; range: ExcelRange }> = [];
    await forEachChunkSync(context, plan.repairs, GOVERNANCE_WRITE_CHUNK, (repair) => {
      const { sheetName, address } = parseCellId(repair.cell);
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(bareAddress(address));
      range.load("formulas");
      readTargets.push({ cell: repair.cell, range });
    });

    const readBack = readTargets.map((t) => ({
      cell: t.cell,
      formula: String(t.range.formulas?.[0]?.[0] ?? ""),
    }));
    const validation = validateRepairedFormulas(readBack);
    if (!validation.ok) {
      throw Object.assign(new Error("formula_repair_incomplete: post-write #REF! remains"), {
        detail: {
          code: "formula_repair_incomplete",
          stillBroken: validation.stillBroken,
          repairs: plan.repairs,
          unresolved: plan.unresolved,
        },
      });
    }
    for (const item of readBack) {
      const planned = plan.repairs.find((r) => r.cell.toLowerCase() === item.cell.toLowerCase());
      if (planned && item.formula !== planned.after) {
        limitations.push(`host normalized formula for ${item.cell}`);
      }
    }

    return {
      scope: input.scope,
      backupId,
      repairs: plan.repairs,
      unresolved: [],
      repairedCount: plan.repairedCount,
      unresolvedCount: 0,
      verified: true,
      limitations,
    };
  });
}

export async function officeJsConvertFormulasToValues(
  input: FormulaConvertToValuesInput,
): Promise<HostResult<FormulaConvertToValuesInfo>> {
  if (!isExcelApi12Supported()) {
    return fail(
      "formula.convertToValues",
      "office-js",
      "ExcelApi 1.2 is not supported in this host",
    );
  }
  requireGovernanceScope(input.scope, input.sheetName, input.range);
  if (input.createBackup === false) {
    return fail(
      "formula.convertToValues",
      "office-js",
      "createBackup=false is not allowed; persistent workbook backup is required",
    );
  }

  return withExcel("formula.convertToValues", async (context) => {
    const limitations: string[] = [];
    const { cells, sourceRange } = await collectFormulaCells(context, input, limitations, {
      includeBackupMetadata: true,
    });
    const formulaCells = cells.filter((c) => c.formula.startsWith("="));
    const backupId = (input.backupId && input.backupId.trim()) || newBackupId();

    if (formulaCells.length === 0) {
      return {
        scope: input.scope,
        backupId,
        convertedFormulaCells: 0,
        verified: true,
        limitations: [...limitations, "no formula cells in scope"],
      };
    }

    await writeFormulaBackup(context, formulaCells, {
      backupId,
      sourceRange,
      create: true,
    });

    // Use values collected with formulas (host calculated snapshot) — no per-cell re-load.
    await forEachChunkSync(context, formulaCells, GOVERNANCE_WRITE_CHUNK, (cell) => {
      const sheet = context.workbook.worksheets.getItem(cell.sheetName);
      const range = sheet.getRange(bareAddress(cell.address));
      range.values = [[(cell.value ?? null) as string | number | boolean | null]];
    });

    const verifyRanges: Array<{ cell: FormulaCellRecord; range: ExcelRange }> = [];
    await forEachChunkSync(context, formulaCells, GOVERNANCE_WRITE_CHUNK, (cell) => {
      const sheet = context.workbook.worksheets.getItem(cell.sheetName);
      const range = sheet.getRange(bareAddress(cell.address));
      range.load("formulas");
      verifyRanges.push({ cell, range });
    });

    let stillFormula = 0;
    for (const item of verifyRanges) {
      if (String(item.range.formulas?.[0]?.[0] ?? "").startsWith("=")) stillFormula += 1;
    }
    if (stillFormula > 0) {
      throw new Error(
        `convert verify failed: ${stillFormula} cell(s) still contain formulas after write`,
      );
    }

    return {
      scope: input.scope,
      backupId,
      convertedFormulaCells: formulaCells.length,
      verified: true,
      limitations,
    };
  });
}
