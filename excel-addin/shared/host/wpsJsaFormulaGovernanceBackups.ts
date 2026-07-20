/**
 * WPS formula backups inspect/restore (strict protocol + removeAfterRestore precheck).
 * removeAfterRestore Clear check runs only after protocol/backupId validation and before any restore writes.
 */
import {
  decodeBackupSheet,
  planRestore,
  summarizeBackups,
  verifyRemainingBackupRows,
} from "../formulaGovernance";
import type {
  FormulaBackupsInspectInfo,
  FormulaBackupsRestoreInfo,
} from "./formulaGovernanceTypes";
import { strictDecodeBackup } from "./officeJsFormulaGovernanceBackup";
import {
  bare,
  canRemoveBackupRows,
  findBackupSheet,
  readBackupMatrix,
  rewriteBackupSheet,
  type WpsRangeExt,
} from "./wpsJsaFormulaGovernanceHelpers";
import { formulaMatrixFrom, getSheet, requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

export async function wpsInspectFormulaBackups(): Promise<HostResult<FormulaBackupsInspectInfo>> {
  const wb = requireWorkbook("formula.backups.inspect");
  if (!wb.ok) return wb;
  try {
    const found = findBackupSheet(wb.data, false);
    if (!found.sheet) {
      return ok({
        backups: [],
        backupCount: 0,
        backupSheetName: null,
        skippedRows: [],
        limitations: found.limitations,
      });
    }
    const decoded = decodeBackupSheet(readBackupMatrix(found.sheet));
    if (!decoded.ok && !decoded.grid) {
      return ok({
        backups: [],
        backupCount: 0,
        backupSheetName: found.sheet.Name,
        skippedRows: decoded.skipped,
        limitations: [...found.limitations, decoded.error ?? "invalid backup"],
        headerError: decoded.error,
      });
    }
    const rows = decoded.grid?.rows ?? [];
    const backups = summarizeBackups(rows);
    const limitations = [...found.limitations];
    if (decoded.error) limitations.push(decoded.error);
    if (decoded.skipped.length) {
      limitations.push(`skipped ${decoded.skipped.length} corrupt/incomplete backup row(s)`);
    }
    return ok({
      backups,
      backupCount: backups.length,
      backupSheetName: found.sheet.Name,
      skippedRows: decoded.skipped,
      limitations,
      headerError: decoded.error,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail("formula.backups.inspect", "wps-jsa", message);
  }
}

export async function wpsRestoreFormulas(input: {
  backupId: string;
  removeAfterRestore?: boolean;
}): Promise<HostResult<FormulaBackupsRestoreInfo>> {
  const wb = requireWorkbook("formula.backups.restore");
  if (!wb.ok) return wb;
  const backupId = input.backupId?.trim();
  if (!backupId) return fail("formula.backups.restore", "wps-jsa", "backupId required");
  try {
    const found = findBackupSheet(wb.data, false);
    if (!found.sheet) {
      return fail("formula.backups.restore", "wps-jsa", "formula_backup_not_found");
    }

    // Strict protocol first — corrupt/not-found must not surface as unsupported Clear.
    const matrix = readBackupMatrix(found.sheet);
    const strict = strictDecodeBackup(matrix);
    if (!strict.ok) {
      return fail("formula.backups.restore", "wps-jsa", strict.error);
    }

    const plan = planRestore(strict.rows, backupId);
    if ("error" in plan) return fail("formula.backups.restore", "wps-jsa", plan.error);

    // Capability precheck after protocol+id validation, before any formula restore.
    if (input.removeAfterRestore === true && !canRemoveBackupRows(found.sheet)) {
      return unsupported(
        "formula.backups.restore",
        "wps-jsa",
        "removeAfterRestore requires UsedRange.Clear to safely rewrite backup sheet without residue",
        "Assumed UsedRange.Clear before full protocol rewrite",
      );
    }

    const restored: Array<{ cell: string; formula: string }> = [];
    const failed: Array<{ cell: string; error: string }> = [];
    const limitations = [
      ...found.limitations,
      "restores formula + numberFormat + locked when available; formulaR1C1/spillAddress backup metadata only",
    ];

    for (const item of plan.items) {
      const cellId = `${item.sheet}!${item.address}`;
      try {
        const sheet = getSheet(wb.data, item.sheet);
        if (!sheet?.Range) throw new Error("sheet missing");
        const range = sheet.Range(bare(item.address)) as WpsRangeExt;
        range.Formula = item.formula;
        if (item.numberFormat) range.NumberFormat = item.numberFormat;
        if (typeof range.Locked === "boolean") range.Locked = item.locked;
        const actual = String(formulaMatrixFrom(range.Formula)?.[0]?.[0] ?? "");
        if (!actual.startsWith("=")) failed.push({ cell: cellId, error: "post-write formula missing" });
        else restored.push({ cell: cellId, formula: actual });
      } catch (error) {
        failed.push({
          cell: cellId,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
    if (failed.length) {
      return fail(
        "formula.backups.restore",
        "wps-jsa",
        "formula_restore_incomplete",
        JSON.stringify({ backupId, restored, failed }),
      );
    }

    if (input.removeAfterRestore === true) {
      const remaining = strict.rows.filter((r) => r.backupId !== backupId);
      rewriteBackupSheet(found.sheet, remaining);
      const after = strictDecodeBackup(readBackupMatrix(found.sheet));
      if (!after.ok) {
        return fail(
          "formula.backups.restore",
          "wps-jsa",
          `removeAfterRestore verify failed: ${after.error}`,
        );
      }
      const multisetError = verifyRemainingBackupRows(remaining, after.rows, backupId);
      if (multisetError) {
        return fail("formula.backups.restore", "wps-jsa", multisetError);
      }
      limitations.push("removed restored backup rows (removeAfterRestore=true)");
    } else {
      limitations.push("backup rows retained (removeAfterRestore default false)");
    }

    return ok({
      backupId,
      restored,
      restoredCount: restored.length,
      failed: [],
      verified: true,
      limitations,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/unavailable|not a function|missing/i.test(message)) {
      return unsupported(
        "formula.backups.restore",
        "wps-jsa",
        message,
        "Assumed Range.Formula + backup sheet",
      );
    }
    return fail("formula.backups.restore", "wps-jsa", message);
  }
}
