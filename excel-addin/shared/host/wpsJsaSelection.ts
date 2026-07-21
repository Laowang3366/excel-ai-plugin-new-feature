import type { HostResult, SelectionInfo } from "./types";
import { ok, unsupported } from "./types";
import { readWpsAddress } from "./wpsJsaAddress";
import { formulaMatrixFrom, matrixFrom, requireApp } from "./wpsJsaRuntime";

/** selection.get for WPS JSA — Address may be string property or zero-arg method. */
export async function wpsGetSelection(): Promise<HostResult<SelectionInfo>> {
  const appResult = requireApp("selection.get");
  if (!appResult.ok) return appResult;
  const selection = appResult.data.Selection;
  const sheet = selection?.Worksheet ?? appResult.data.ActiveWorkbook?.ActiveSheet;
  if (!selection || !sheet) {
    return unsupported(
      "selection.get",
      "wps-jsa",
      "Selection or Worksheet unavailable",
      "Assumed Application.Selection / Worksheet",
    );
  }
  const address = readWpsAddress(selection);
  if (!address) {
    return unsupported(
      "selection.get",
      "wps-jsa",
      "Selection.Address unavailable or not a usable string/method result",
      "Assumed Application.Selection.Address (property or zero-arg method)",
    );
  }
  return ok({
    sheetName: sheet.Name,
    address,
    values: matrixFrom(selection.Value2),
    formulas: formulaMatrixFrom(selection.Formula),
  });
}
