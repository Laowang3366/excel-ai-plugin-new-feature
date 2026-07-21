import { requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult, SheetInfo, WorkbookInspectInfo } from "./types";
import { ok, unsupported } from "./types";
import { readWpsAddress } from "./wpsJsaAddress";

/** WPS workbook.inspect (active used range only; per-sheet dims not verified). */
export async function wpsInspectWorkbook(
  listSheets: () => Promise<HostResult<SheetInfo[]>>,
): Promise<HostResult<WorkbookInspectInfo>> {
  const workbookResult = requireWorkbook("workbook.inspect");
  if (!workbookResult.ok) return workbookResult;
  const workbook = workbookResult.data;
  const sheetsResult = await listSheets();
  if (!sheetsResult.ok) return sheetsResult;
  const active = workbook.ActiveSheet;
  if (!active?.Name) {
    return unsupported(
      "workbook.inspect",
      "wps-jsa",
      "ActiveSheet is unavailable",
      "Requires ActiveWorkbook.ActiveSheet.Name",
    );
  }
  const usedAddress = readWpsAddress(active.UsedRange) ?? null;
  return ok({
    workbookName: workbook.Name,
    activeSheetName: active.Name,
    sheetCount: sheetsResult.data.length,
    usedRangeAddress: usedAddress,
    sheets: sheetsResult.data,
  });
}
