/**
 * Pure TypeScript formula governance core (phase A).
 * No HostAdapter / tool registration — stage B wires Office.js / WPS.
 */

export {
  makeCellId,
  normalizeA1Address,
  parseCellId,
  isA1Like,
  unescapeSheetName,
} from "./address";

export {
  removeStringLiterals,
  parseFormulaReferences,
  referencesToEdges,
  type ParsedReference,
} from "./references";

export { buildDependencyReport, findCycles } from "./dependencyGraph";

export {
  planFormulaRepairs,
  validateRepairedFormulas,
  type RepairOptions,
} from "./repair";

export {
  FORMULA_BACKUP_MAGIC,
  FORMULA_BACKUP_HEADERS,
  DEPENDENCY_LIMITATIONS,
  type FormulaCellRecord,
  type FormulaEdge,
  type FormulaEdgeKind,
  type FormulaKind,
  type FormulaDependencyNode,
  type FormulaDependencyReport,
  type BrokenReference,
  type FormulaReplacement,
  type FormulaRepairItem,
  type FormulaRepairPlan,
  type FormulaBackupRow,
  type FormulaBackupSummary,
  type FormulaBackupHeader,
  type FormulaRestorePlan,
  type FormulaRestorePlanItem,
} from "./types";

export {
  isBackupMagic,
  createBackupRows,
  encodeBackupSheet,
  decodeBackupSheet,
  summarizeBackups,
  planRestore,
  type BackupSheetGrid,
} from "./backup";

export {
  isFormula,
  leadingFunction,
  isDynamicArray,
  classifyFormula,
  tryClassifyFormula,
} from "./classification";
