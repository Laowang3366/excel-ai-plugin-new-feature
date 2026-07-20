/**
 * Host-facing formula governance types (dependencies / repair / convert / backups).
 * Pure core lives in shared/formulaGovernance; this is the HostAdapter contract.
 */

import type {
  FormulaBackupSummary,
  FormulaDependencyReport,
  FormulaRepairItem,
  FormulaReplacement,
} from "../formulaGovernance";

export type FormulaGovernanceScope = "workbook" | "sheet" | "target";

export interface FormulaGovernanceScopedInput {
  scope: FormulaGovernanceScope;
  sheetName?: string;
  range?: string;
}

export type FormulaDependenciesInspectInput = FormulaGovernanceScopedInput;

export interface FormulaDependenciesInspectInfo {
  scope: FormulaGovernanceScope;
  report: FormulaDependencyReport;
  /** Always includes text-parse limitations from pure core. */
  limitations: string[];
}

export interface FormulaReferencesRepairInput extends FormulaGovernanceScopedInput {
  replacements: FormulaReplacement[];
  applyAllMappings?: boolean;
}

export interface FormulaReferencesRepairInfo {
  scope: FormulaGovernanceScope;
  backupId: string;
  repairs: FormulaRepairItem[];
  unresolved: FormulaRepairItem[];
  repairedCount: number;
  unresolvedCount: number;
  verified: boolean;
  limitations: string[];
}

export interface FormulaConvertToValuesInput extends FormulaGovernanceScopedInput {
  /** Always true in this add-in — persistent workbook backup is required. */
  createBackup?: boolean;
  backupId?: string;
}

export interface FormulaConvertToValuesInfo {
  scope: FormulaGovernanceScope;
  backupId: string;
  convertedFormulaCells: number;
  verified: boolean;
  limitations: string[];
}

export interface FormulaBackupsInspectInfo {
  backups: FormulaBackupSummary[];
  backupCount: number;
  backupSheetName: string | null;
  skippedRows: number[];
  limitations: string[];
  /** Present when magic/header is invalid but partial parse may still apply. */
  headerError?: string;
}

export interface FormulaBackupsRestoreInput {
  backupId: string;
  /** Default false — never auto-delete backup rows. */
  removeAfterRestore?: boolean;
}

export interface FormulaRestoreCellResult {
  cell: string;
  formula: string;
}

export interface FormulaRestoreFailedCell {
  cell: string;
  error: string;
}

export interface FormulaBackupsRestoreInfo {
  backupId: string;
  restored: FormulaRestoreCellResult[];
  restoredCount: number;
  failed: FormulaRestoreFailedCell[];
  verified: boolean;
  limitations: string[];
}

export const FORMULA_BACKUP_SHEET_PREFIX = "_WenggeFormulaBackup";
export const MAX_GOVERNANCE_FORMULA_CELLS = 5_000;
