/** Pure formula-governance types (no host I/O). */

export type FormulaEdgeKind = "same-sheet" | "cross-sheet" | "external";

export type FormulaKind = "plain" | "dynamic" | "legacyArray";

/** One formula-bearing cell supplied by the host layer. */
export interface FormulaCellRecord {
  sheetName: string;
  /** A1 address without sheet, e.g. A1 or $B$2:$C$3 */
  address: string;
  formula: string;
  value?: unknown;
  formulaR1C1?: string;
  numberFormat?: string;
  locked?: boolean;
  spillAddress?: string;
}

export interface FormulaEdge {
  from: string;
  to: string;
  kind: FormulaEdgeKind;
  reference: string;
}

export interface FormulaDependencyNode {
  id: string;
  sheet: string;
  address: string;
  formula: string;
  value?: unknown;
  precedents: string[];
  dependents: string[];
}

export interface BrokenReference {
  cell: string;
  formula: string;
  reason: string;
}

export interface FormulaDependencyReport {
  nodes: FormulaDependencyNode[];
  edges: FormulaEdge[];
  cycles: { path: string[] }[];
  brokenReferences: BrokenReference[];
  formulaCount: number;
  edgeCount: number;
  /** Always present: this graph is text-parse only. */
  limitations: readonly string[];
}

export interface FormulaReplacement {
  find: string;
  replace: string;
}

export interface FormulaRepairItem {
  cell: string;
  before: string;
  after: string;
  strategy: "mapping";
}

export interface FormulaRepairPlan {
  repairs: FormulaRepairItem[];
  unresolved: FormulaRepairItem[];
  repairedCount: number;
  unresolvedCount: number;
  /** True only when every repaired formula no longer contains #REF! */
  complete: boolean;
}

export const FORMULA_BACKUP_MAGIC = "WENGGE_FORMULA_BACKUP_V1";
export const FORMULA_BACKUP_HEADERS = [
  "backupId",
  "createdAt",
  "sheet",
  "address",
  "formula",
  "formulaR1C1",
  "numberFormat",
  "locked",
  "spillAddress",
  "sourceRange",
] as const;

export type FormulaBackupHeader = (typeof FORMULA_BACKUP_HEADERS)[number];

/** One persisted backup row (workbook sheet storage, not session memory). */
export interface FormulaBackupRow {
  backupId: string;
  createdAt: string;
  sheet: string;
  address: string;
  formula: string;
  formulaR1C1: string;
  numberFormat: string;
  locked: boolean;
  spillAddress: string;
  sourceRange: string;
}

export interface FormulaBackupSummary {
  backupId: string;
  createdAt: string;
  formulaCount: number;
  sheets: string[];
  sourceRanges: string[];
}

export interface FormulaRestorePlanItem {
  sheet: string;
  address: string;
  formula: string;
  formulaR1C1: string;
  numberFormat: string;
  locked: boolean;
  spillAddress: string;
  sourceRange: string;
}

export interface FormulaRestorePlan {
  backupId: string;
  items: FormulaRestorePlanItem[];
}

export const DEPENDENCY_LIMITATIONS = [
  "text-parse-only",
  "no-excel-engine-circularReference",
  "no-structured-table-refs",
  "no-indirect-or-offset-resolution",
  "no-3d-sheet-range-refs",
  "no-defined-name-expansion",
  "string-literals-skipped",
] as const;
