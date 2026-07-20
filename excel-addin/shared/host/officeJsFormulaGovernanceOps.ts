/**
 * Office.js formula governance: dependencies inspect, references repair, convertToValues.
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
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail } from "./types";

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
    const { cells } = await collectFormulaCells(context, input, limitations);
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
    const { cells, sourceRange } = await collectFormulaCells(context, input, limitations);
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

    for (const repair of plan.repairs) {
      const bang = repair.cell.lastIndexOf("!");
      const sheetName = repair.cell.slice(0, bang);
      const address = repair.cell.slice(bang + 1);
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(bareAddress(address));
      range.formulas = [[repair.after]];
      await context.sync();
    }

    const readBack: Array<{ cell: string; formula: string }> = [];
    for (const repair of plan.repairs) {
      const bang = repair.cell.lastIndexOf("!");
      const sheetName = repair.cell.slice(0, bang);
      const address = repair.cell.slice(bang + 1);
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(bareAddress(address));
      range.load("formulas");
      await context.sync();
      readBack.push({ cell: repair.cell, formula: String(range.formulas?.[0]?.[0] ?? "") });
    }
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
    const { cells, sourceRange } = await collectFormulaCells(context, input, limitations);
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

    let converted = 0;
    for (const cell of formulaCells) {
      const sheet = context.workbook.worksheets.getItem(cell.sheetName);
      const range = sheet.getRange(bareAddress(cell.address));
      range.load("values");
      await context.sync();
      const value = range.values?.[0]?.[0] ?? null;
      range.values = [[value as string | number | boolean | null]];
      await context.sync();
      converted += 1;
    }

    let stillFormula = 0;
    for (const cell of formulaCells) {
      const sheet = context.workbook.worksheets.getItem(cell.sheetName);
      const range = sheet.getRange(bareAddress(cell.address));
      range.load("formulas");
      await context.sync();
      if (String(range.formulas?.[0]?.[0] ?? "").startsWith("=")) stillFormula += 1;
    }
    if (stillFormula > 0) {
      throw new Error(
        `convert verify failed: ${stillFormula} cell(s) still contain formulas after write`,
      );
    }

    return {
      scope: input.scope,
      backupId,
      convertedFormulaCells: converted,
      verified: true,
      limitations,
    };
  });
}
