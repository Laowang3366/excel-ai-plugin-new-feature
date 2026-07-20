/**
 * Formula protection via Range.formulas + Range.format.protection.locked (ExcelApi 1.2)
 * and Worksheet.protection (existing). Does not whole-sheet lock as a substitute for
 * per-formula locking.
 */
export { officeJsInspectFormulaProtection } from "./officeJsFormulaProtectionInspect";
export { officeJsManageFormulaProtection } from "./officeJsFormulaProtectionManage";
