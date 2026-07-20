/** Formula cell lock / sheet protect governance (Office.js subset). */

export type FormulaProtectionScope = "workbook" | "sheet" | "target";
export type FormulaProtectionCommand = "lock" | "unlock";

export interface FormulaProtectionInspectInput {
  scope: FormulaProtectionScope;
  /** Required for sheet|target; ignored for workbook (all sheets scanned). */
  sheetName?: string;
  /** Required for target; optional for sheet (defaults to used range). */
  range?: string;
}

export interface FormulaProtectionManageInput {
  command: FormulaProtectionCommand;
  scope: FormulaProtectionScope;
  sheetName?: string;
  range?: string;
  /**
   * Request-scoped only — never persisted, logged, or returned in HostResult.
   * Used for temporary unprotect and optional protectSheet.
   */
  password?: string;
  /**
   * lock only (default true): unlock entire target range first so non-formula
   * input cells remain editable after sheet protect. Does NOT unlock outside range.
   */
  unlockInputs?: boolean;
  /** lock only (default true): protect sheet after locking formula cells. */
  protectSheet?: boolean;
}

export interface FormulaProtectionSheetSummary {
  sheetName: string;
  /** Bare or sheet-qualified host address of the scanned range. */
  address: string;
  formulaCount: number;
  lockedFormulaCount: number;
  sheetProtected: boolean;
  /** Human-readable limits for this sheet (chunking, mixed lock, API). */
  limitations: string[];
}

export interface FormulaProtectionInspectInfo {
  scope: FormulaProtectionScope;
  sheets: FormulaProtectionSheetSummary[];
  formulaCount: number;
  lockedFormulaCount: number;
  limitations: string[];
}

export interface FormulaProtectionManageInfo {
  command: FormulaProtectionCommand;
  scope: FormulaProtectionScope;
  unlockInputs: boolean;
  protectSheet: boolean;
  /** Post-write verification via re-inspect (no password fields). */
  protection: FormulaProtectionInspectInfo;
  verified: boolean;
  limitations: string[];
}
