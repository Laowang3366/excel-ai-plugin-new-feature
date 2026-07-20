/**
 * Office.js formula backup inspect / restore (WENGGE_FORMULA_BACKUP_V1).
 */
import { planRestore } from "../formulaGovernance";
import type {
  FormulaBackupsInspectInfo,
  FormulaBackupsRestoreInfo,
} from "./formulaGovernanceTypes";
import {
  loadBackupRows,
  rewriteBackupSheet,
  inspectBackupSheet,
} from "./officeJsFormulaGovernanceBackup";
import { bareAddress } from "./officeJsFormulaGovernanceCollect";
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail } from "./types";

export async function officeJsInspectFormulaBackups(): Promise<
  HostResult<FormulaBackupsInspectInfo>
> {
  return withExcel("formula.backups.inspect", async (context) => inspectBackupSheet(context));
}

export async function officeJsRestoreFormulas(input: {
  backupId: string;
  removeAfterRestore?: boolean;
}): Promise<HostResult<FormulaBackupsRestoreInfo>> {
  const backupId = input.backupId?.trim();
  if (!backupId) {
    return fail("formula.backups.restore", "office-js", "backupId required");
  }

  return withExcel("formula.backups.restore", async (context) => {
    const loaded = await loadBackupRows(context);
    if (!loaded.sheet || loaded.error === "formula_backup_not_found") {
      throw Object.assign(new Error("formula_backup_not_found"), {
        detail: { code: "formula_backup_not_found" },
      });
    }
    if (loaded.rows.length === 0 && loaded.error) {
      throw Object.assign(new Error(loaded.error), {
        detail: { code: loaded.error },
      });
    }

    const plan = planRestore(loaded.rows, backupId);
    if ("error" in plan) {
      throw Object.assign(new Error(plan.error), {
        detail: { code: plan.error, backupId },
      });
    }

    const restored: Array<{ cell: string; formula: string }> = [];
    const failed: Array<{ cell: string; error: string }> = [];
    const limitations = [...loaded.limitations];

    for (const item of plan.items) {
      const cellId = `${item.sheet}!${item.address}`;
      try {
        const sheet = context.workbook.worksheets.getItem(item.sheet);
        const range = sheet.getRange(bareAddress(item.address));
        range.formulas = [[item.formula]];
        if (item.numberFormat) {
          range.numberFormat = [[item.numberFormat]];
        }
        await context.sync();
        if (item.locked === true || item.locked === false) {
          try {
            if (range.format?.protection) {
              range.format.protection.locked = item.locked;
              await context.sync();
            }
          } catch {
            limitations.push(`locked restore skipped for ${cellId}`);
          }
        }
        range.load("formulas");
        await context.sync();
        const actual = String(range.formulas?.[0]?.[0] ?? "");
        if (!actual.startsWith("=")) {
          failed.push({ cell: cellId, error: "post-write formula missing" });
        } else {
          restored.push({ cell: cellId, formula: actual });
        }
      } catch (error) {
        failed.push({
          cell: cellId,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    if (failed.length > 0) {
      throw Object.assign(new Error("formula_restore_incomplete"), {
        detail: { code: "formula_restore_incomplete", backupId, restored, failed },
      });
    }

    if (input.removeAfterRestore === true) {
      const remaining = loaded.rows.filter((r) => r.backupId !== backupId);
      await rewriteBackupSheet(context, loaded.sheet, remaining);
      limitations.push("removed restored backup rows (removeAfterRestore=true)");
    } else {
      limitations.push("backup rows retained (removeAfterRestore default false)");
    }

    return {
      backupId,
      restored,
      restoredCount: restored.length,
      failed: [],
      verified: true,
      limitations,
    };
  });
}
