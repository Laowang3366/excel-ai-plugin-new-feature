/**
 * Office.js formula protection inspect.
 */
import type {
  FormulaProtectionInspectInfo,
  FormulaProtectionInspectInput,
} from "./formulaProtectionTypes";
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { unsupported } from "./types";
import {
  inspectAll,
  isExcelApi12Supported,
  LOCKED_EVIDENCE,
} from "./officeJsFormulaProtectionHelpers";

export async function officeJsInspectFormulaProtection(
  input: FormulaProtectionInspectInput,
): Promise<HostResult<FormulaProtectionInspectInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "formula.protection.inspect",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      LOCKED_EVIDENCE,
    );
  }
  return withExcel("formula.protection.inspect", async (context) => inspectAll(context, input));
}

