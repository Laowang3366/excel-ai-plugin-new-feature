import type {
  ChartInfo,
  ChartType,
  HostResult,
  TableInfo,
  TableUnlistInfo,
} from "./types";
import { unsupported } from "./types";
export { wpsReadFormat, wpsWriteFormat } from "./wpsJsaFormat";

const EVIDENCE =
  "In-repo bridge only covers Application/ActiveWorkbook/JSIDE CodeModule; no verified format/table/chart contract";

/** Phase3 WPS capabilities without in-repo evidence → typed unsupported. */
export async function wpsListTables(_sheetName?: string): Promise<HostResult<TableInfo[]>> {
  return unsupported(
    "table.list",
    "wps-jsa",
    "ListObjects are not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsCreateTable(_input: {
  sheetName: string;
  address: string;
  name?: string;
  hasHeaders?: boolean;
}): Promise<HostResult<TableInfo>> {
  return unsupported(
    "table.create",
    "wps-jsa",
    "ListObjects.Add is not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsDeleteTable(
  _sheetName: string,
  _tableName: string,
): Promise<HostResult<{ deleted: string }>> {
  return unsupported(
    "table.delete",
    "wps-jsa",
    "ListObject.Delete is not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsUnlistTable(
  _sheetName: string,
  _tableName: string,
): Promise<HostResult<TableUnlistInfo>> {
  return unsupported(
    "table.unlist",
    "wps-jsa",
    "ListObjects/Unlist/convertToRange are not verified in this repository for WPS JSA",
    "No in-repo WPS JSA ListObjects/Unlist/convertToRange contract",
  );
}

export async function wpsListCharts(_sheetName?: string): Promise<HostResult<ChartInfo[]>> {
  return unsupported(
    "chart.list",
    "wps-jsa",
    "ChartObjects are not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsCreateChart(_input: {
  sheetName: string;
  sourceRange: string;
  chartType?: ChartType;
  name?: string;
  title?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}): Promise<HostResult<ChartInfo>> {
  return unsupported(
    "chart.create",
    "wps-jsa",
    "Chart create/source binding is not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsDeleteChart(
  _sheetName: string,
  _chartName: string,
): Promise<HostResult<{ deleted: string }>> {
  return unsupported(
    "chart.delete",
    "wps-jsa",
    "ChartObject.Delete is not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsUpdateTable(_input: unknown) {
  return unsupported(
    "table.update",
    "wps-jsa",
    "Table update is not verified for WPS JSA",
    "No in-repo WPS ListObject update contract",
  ) as HostResult<never>;
}

export async function wpsGetTableFilter(_input: unknown) {
  return unsupported(
    "table.filter.get",
    "wps-jsa",
    "Table filter is not verified for WPS JSA",
    "No in-repo WPS ListObject.AutoFilter contract",
  ) as HostResult<never>;
}

export async function wpsApplyTableFilter(_input: unknown) {
  return unsupported(
    "table.filter.apply",
    "wps-jsa",
    "Table filter is not verified for WPS JSA",
    "No in-repo WPS ListObject.AutoFilter contract",
  ) as HostResult<never>;
}

export async function wpsClearTableFilter(_input: unknown) {
  return unsupported(
    "table.filter.clear",
    "wps-jsa",
    "Table filter is not verified for WPS JSA",
    "No in-repo WPS ListObject.AutoFilter contract",
  ) as HostResult<never>;
}

export async function wpsGetTableSort(_input: unknown) {
  return unsupported(
    "table.sort.get",
    "wps-jsa",
    "Table sort is not verified for WPS JSA",
    "No in-repo WPS ListObject.Sort contract",
  ) as HostResult<never>;
}

export async function wpsApplyTableSort(_input: unknown) {
  return unsupported(
    "table.sort.apply",
    "wps-jsa",
    "Table sort is not verified for WPS JSA",
    "No in-repo WPS ListObject.Sort contract",
  ) as HostResult<never>;
}

export async function wpsClearTableSort(_input: unknown) {
  return unsupported(
    "table.sort.clear",
    "wps-jsa",
    "Table sort is not verified for WPS JSA",
    "No in-repo WPS ListObject.Sort contract",
  ) as HostResult<never>;
}

export async function wpsUpdateChart(_input: unknown) {
  return unsupported(
    "chart.update",
    "wps-jsa",
    "Chart update is not verified for WPS JSA",
    "No in-repo WPS ChartObject update contract",
  ) as HostResult<never>;
}

export {
  wpsGetSheetVisibility,
  wpsSetSheetVisibility,
} from "./wpsJsaSheetVisibility";
export {
  wpsGetSheetProtection,
  wpsProtectSheet,
  wpsUnprotectSheet,
} from "./wpsJsaSheetProtection";
export {
  wpsListNamedRanges,
  wpsCreateNamedRange,
  wpsUpdateNamedRange,
  wpsDeleteNamedRange,
} from "./wpsJsaNamedRanges";

export async function wpsGetSheetDisplay(_sheetName: string) {
  return unsupported(
    "sheet.display.get",
    "wps-jsa",
    "sheet.display.get is not verified for WPS JSA",
    "No in-repo WPS Worksheet tabColor/showGridlines/showHeadings contract",
  ) as HostResult<never>;
}

export async function wpsSetSheetDisplay(_input: unknown) {
  return unsupported(
    "sheet.display.set",
    "wps-jsa",
    "sheet.display.set is not verified for WPS JSA",
    "No in-repo WPS Worksheet tabColor/showGridlines/showHeadings contract",
  ) as HostResult<never>;
}

export async function wpsGetSheetFreeze(_sheetName: string) {
  return unsupported(
    "sheet.freeze.get",
    "wps-jsa",
    "sheet.freeze.get is not verified for WPS JSA",
    "No in-repo WPS FreezePanes contract",
  ) as HostResult<never>;
}

export async function wpsSetSheetFreeze(_input: unknown) {
  return unsupported(
    "sheet.freeze.set",
    "wps-jsa",
    "sheet.freeze.set is not verified for WPS JSA",
    "No in-repo WPS FreezePanes contract",
  ) as HostResult<never>;
}

export async function wpsGetSheetPageLayout(_sheetName: string) {
  return unsupported(
    "sheet.pageLayout.get",
    "wps-jsa",
    "sheet.pageLayout.get is not verified for WPS JSA",
    "No in-repo WPS PageLayout/print settings contract",
  ) as HostResult<never>;
}

export async function wpsSetSheetPageLayout(_input: unknown) {
  return unsupported(
    "sheet.pageLayout.set",
    "wps-jsa",
    "sheet.pageLayout.set is not verified for WPS JSA",
    "No in-repo WPS PageLayout/print settings contract",
  ) as HostResult<never>;
}

const SHAPE_EVIDENCE = "No in-repo WPS Shapes/ShapeObjects contract";

export async function wpsListShapes(_sheetName?: string) {
  return unsupported(
    "shape.list",
    "wps-jsa",
    "shape.list is not verified for WPS JSA",
    SHAPE_EVIDENCE,
  ) as HostResult<never>;
}
export async function wpsCreateShape(_input: unknown) {
  return unsupported(
    "shape.create",
    "wps-jsa",
    "shape.create is not verified for WPS JSA",
    SHAPE_EVIDENCE,
  ) as HostResult<never>;
}
export async function wpsDeleteShape(_sheetName: string, _shapeName: string) {
  return unsupported(
    "shape.delete",
    "wps-jsa",
    "shape.delete is not verified for WPS JSA",
    SHAPE_EVIDENCE,
  ) as HostResult<never>;
}
export async function wpsUpdateShape(_input: unknown) {
  return unsupported(
    "shape.update",
    "wps-jsa",
    "shape.update is not verified for WPS JSA",
    SHAPE_EVIDENCE,
  ) as HostResult<never>;
}

export const wpsShapeSurface = {
  listShapes: wpsListShapes,
  createShape: wpsCreateShape,
  deleteShape: wpsDeleteShape,
  updateShape: wpsUpdateShape,
};

export async function wpsInspectFormulaProtection(_input: unknown) {
  return unsupported(
    "formula.protection.inspect",
    "wps-jsa",
    "Formula protection inspect is not verified for WPS JSA",
    "No in-repo WPS Range.Locked / ProtectContents contract — do not guess API",
  ) as HostResult<never>;
}

export async function wpsManageFormulaProtection(_input: unknown) {
  return unsupported(
    "formula.protection.manage",
    "wps-jsa",
    "Formula protection manage is not verified for WPS JSA",
    "No in-repo WPS Range.Locked / ProtectContents contract — do not guess API",
  ) as HostResult<never>;
}
