/**
 * Office.js runtime detection and Excel.run batch helpers.
 * Excel surface types live in officeJsExcelTypes.ts.
 */
export * from "./officeJsExcelTypes";
export { normalizeMatrix, normalizeFormulas, firstNumberFormat } from "./officeJsNormalize";
export { loadRangeFormat, readFormatFromRange } from "./officeJsRangeFormat";

import type { ExcelRequestContext, ExcelRunFn } from "./officeJsExcelTypes";
import type { HostResult, HostRuntimeCapabilities } from "./types";
import { fail, ok, unsupported } from "./types";

export function getExcelRun(): ExcelRunFn | null {
  if (typeof window === "undefined") return null;
  return window.Excel?.run ?? null;
}

export function getOfficeJsRuntimeCapabilities(): HostRuntimeCapabilities {
  if (typeof window === "undefined") {
    return { dynamicArrayFunctionsEnabled: false };
  }
  const isSetSupported = window.Office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") {
    return { dynamicArrayFunctionsEnabled: false };
  }
  let dynamicArrayFunctionsEnabled = false;
  try {
    dynamicArrayFunctionsEnabled = Boolean(
      isSetSupported.call(
        window.Office?.context?.requirements,
        "ExcelApi",
        "1.12",
      ),
    );
  } catch {
    // Requirement-set detection is advisory and must remain fail-safe.
  }
  return {
    dynamicArrayFunctionsEnabled,
  };
}


export async function withExcel<T>(
  capability: string,
  fn: (context: ExcelRequestContext) => Promise<T>,
): Promise<HostResult<T>> {
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      capability,
      "office-js",
      "Excel.run is not available in this runtime",
      "Requires Microsoft Office Excel with Office.js loaded",
    );
  }
  try {
    return ok(await run(fn));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    // Precheck already passed when Excel.run exists: batch/load/sync errors are ordinary failures.
    return fail(capability, "office-js", message);
  }
}
