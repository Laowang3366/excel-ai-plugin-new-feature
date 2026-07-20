/**
 * WPS JSA formula governance operations (member-probed).
 * removeAfterRestore=true requires UsedRange.Clear precheck; otherwise typed unsupported
 * before any restore/delete.
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
import {
  appendBackup,
  bare,
  collectAll,
  ensureVeryHidden,
  findBackupSheet,
  newId,
  requireScope,
} from "./wpsJsaFormulaGovernanceHelpers";
import { formulaMatrixFrom, getSheet, matrixFrom, requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

export async function wpsInspectFormulaDependencies(
  input: FormulaDependenciesInspectInput,
): Promise<HostResult<FormulaDependenciesInspectInfo>> {
  const wb = requireWorkbook("formula.dependencies.inspect");
  if (!wb.ok) return wb;
  try {
    requireScope(input.scope, input.sheetName, input.range);
    const limitations: string[] = [];
    const { cells } = collectAll(wb.data, input, limitations);
    const report = buildDependencyReport(cells);
    return ok({
      scope: input.scope,
      report,
      limitations: [...limitations, ...report.limitations],
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/unavailable|not a function|missing/i.test(message)) {
      return unsupported(
        "formula.dependencies.inspect",
        "wps-jsa",
        message,
        "Assumed Worksheets/Range.Formula/UsedRange with member probe",
      );
    }
    return fail("formula.dependencies.inspect", "wps-jsa", message);
  }
}

export async function wpsRepairFormulaReferences(
  input: FormulaReferencesRepairInput,
): Promise<HostResult<FormulaReferencesRepairInfo>> {
  const wb = requireWorkbook("formula.references.repair");
  if (!wb.ok) return wb;
  try {
    requireScope(input.scope, input.sheetName, input.range);
    if (!Array.isArray(input.replacements) || input.replacements.length === 0) {
      return fail("formula.references.repair", "wps-jsa", "replacements required");
    }
    const limitations: string[] = [];
    const { cells, sourceRange } = collectAll(wb.data, input, limitations);
    const plan = planFormulaRepairs(cells, input.replacements, {
      applyAllMappings: input.applyAllMappings === true,
    });
    if (!plan.complete) {
      return fail(
        "formula.references.repair",
        "wps-jsa",
        "formula_repair_incomplete",
        JSON.stringify({
          code: "formula_repair_incomplete",
          repairs: plan.repairs,
          unresolved: plan.unresolved,
        }),
      );
    }
    if (plan.repairs.length === 0) {
      return ok({
        scope: input.scope,
        backupId: "",
        repairs: [],
        unresolved: [],
        repairedCount: 0,
        unresolvedCount: 0,
        verified: true,
        limitations: [...limitations, "no matching formula cells to repair"],
      });
    }
    const backupId = newId();
    // Backup sheet create+hide BEFORE mutating formulas
    let found;
    try {
      found = findBackupSheet(wb.data, true);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (/visibility/i.test(message)) {
        return fail("formula.references.repair", "wps-jsa", message);
      }
      if (/Add unavailable/i.test(message)) {
        return unsupported(
          "formula.references.repair",
          "wps-jsa",
          message,
          "Assumed Worksheets.Add",
        );
      }
      throw error;
    }
    limitations.push(...found.limitations);
    if (!found.sheet) {
      return unsupported(
        "formula.references.repair",
        "wps-jsa",
        "cannot create backup sheet",
        "Assumed Worksheets.Add",
      );
    }
    ensureVeryHidden(found.sheet);
    const backupCells = plan.repairs
      .map((r) =>
        cells.find((c) => `${c.sheetName}!${c.address}`.toLowerCase() === r.cell.toLowerCase()),
      )
      .filter((c): c is FormulaCellRecord => Boolean(c));
    appendBackup(found.sheet, backupCells, backupId, sourceRange);

    for (const repair of plan.repairs) {
      const bang = repair.cell.lastIndexOf("!");
      const sheet = getSheet(wb.data, repair.cell.slice(0, bang));
      if (!sheet?.Range) throw new Error(`sheet missing for ${repair.cell}`);
      sheet.Range(bare(repair.cell.slice(bang + 1))).Formula = repair.after;
    }

    const readBack: Array<{ cell: string; formula: string }> = [];
    for (const repair of plan.repairs) {
      const bang = repair.cell.lastIndexOf("!");
      const sheet = getSheet(wb.data, repair.cell.slice(0, bang));
      const f = String(
        formulaMatrixFrom(sheet!.Range(bare(repair.cell.slice(bang + 1))).Formula)?.[0]?.[0] ?? "",
      );
      readBack.push({ cell: repair.cell, formula: f });
    }
    const validation = validateRepairedFormulas(readBack);
    if (!validation.ok) {
      return fail(
        "formula.references.repair",
        "wps-jsa",
        "formula_repair_incomplete",
        JSON.stringify(validation),
      );
    }
    return ok({
      scope: input.scope,
      backupId,
      repairs: plan.repairs,
      unresolved: [],
      repairedCount: plan.repairedCount,
      unresolvedCount: 0,
      verified: true,
      limitations,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/visibility/i.test(message)) {
      return fail("formula.references.repair", "wps-jsa", message);
    }
    if (/unavailable|not a function|missing/i.test(message)) {
      return unsupported(
        "formula.references.repair",
        "wps-jsa",
        message,
        "Assumed Range.Formula + Worksheets.Add",
      );
    }
    return fail("formula.references.repair", "wps-jsa", message);
  }
}

export async function wpsConvertFormulasToValues(
  input: FormulaConvertToValuesInput,
): Promise<HostResult<FormulaConvertToValuesInfo>> {
  const wb = requireWorkbook("formula.convertToValues");
  if (!wb.ok) return wb;
  try {
    requireScope(input.scope, input.sheetName, input.range);
    if (input.createBackup === false) {
      return fail(
        "formula.convertToValues",
        "wps-jsa",
        "createBackup=false is not allowed; persistent workbook backup is required",
      );
    }
    const limitations: string[] = [];
    const { cells, sourceRange } = collectAll(wb.data, input, limitations);
    const formulaCells = cells.filter((c) => c.formula.startsWith("="));
    const backupId = input.backupId?.trim() || newId();
    if (formulaCells.length === 0) {
      return ok({
        scope: input.scope,
        backupId,
        convertedFormulaCells: 0,
        verified: true,
        limitations: [...limitations, "no formula cells in scope"],
      });
    }
    let found;
    try {
      found = findBackupSheet(wb.data, true);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (/visibility/i.test(message)) {
        return fail("formula.convertToValues", "wps-jsa", message);
      }
      if (/Add unavailable/i.test(message)) {
        return unsupported("formula.convertToValues", "wps-jsa", message, "Assumed Worksheets.Add");
      }
      throw error;
    }
    limitations.push(...found.limitations);
    if (!found.sheet) {
      return unsupported(
        "formula.convertToValues",
        "wps-jsa",
        "cannot create backup sheet",
        "Assumed Worksheets.Add",
      );
    }
    ensureVeryHidden(found.sheet);
    appendBackup(found.sheet, formulaCells, backupId, sourceRange);

    let converted = 0;
    for (const cell of formulaCells) {
      const sheet = getSheet(wb.data, cell.sheetName);
      if (!sheet?.Range) throw new Error(`sheet missing ${cell.sheetName}`);
      const range = sheet.Range(bare(cell.address));
      const value = matrixFrom(range.Value2)?.[0]?.[0] ?? null;
      range.Value2 = value;
      converted += 1;
    }
    let still = 0;
    for (const cell of formulaCells) {
      const sheet = getSheet(wb.data, cell.sheetName)!;
      const f = String(
        formulaMatrixFrom(sheet.Range(bare(cell.address)).Formula)?.[0]?.[0] ?? "",
      );
      if (f.startsWith("=")) still += 1;
    }
    if (still > 0) {
      return fail(
        "formula.convertToValues",
        "wps-jsa",
        `convert verify failed: ${still} formula(s) remain`,
      );
    }
    return ok({
      scope: input.scope,
      backupId,
      convertedFormulaCells: converted,
      verified: true,
      limitations,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/visibility/i.test(message)) {
      return fail("formula.convertToValues", "wps-jsa", message);
    }
    if (/unavailable|not a function|missing/i.test(message)) {
      return unsupported(
        "formula.convertToValues",
        "wps-jsa",
        message,
        "Assumed Range.Value2/Formula",
      );
    }
    return fail("formula.convertToValues", "wps-jsa", message);
  }
}
