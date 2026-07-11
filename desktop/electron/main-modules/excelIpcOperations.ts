import type { ExcelWorkbookBridge } from "../agent/tools/contracts/excel";

type RangeExpandMode = "none" | "spill" | "currentArray" | "currentRegion";

export async function readExcelRangeForIpc(
  bridge: ExcelWorkbookBridge | null,
  sheetName: string,
  range: string,
  expand?: RangeExpandMode,
) {
  if (!bridge) throw new Error("Excel 桥接未初始化");
  return bridge.readRange(sheetName, range, expand);
}

export async function inspectExcelWorkbookForIpc(bridge: ExcelWorkbookBridge | null) {
  if (!bridge) throw new Error("Excel 桥接未初始化");
  return bridge.inspectWorkbook();
}
