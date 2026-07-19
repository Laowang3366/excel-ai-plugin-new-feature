import type {
  ChartInfo,
  ChartType,
  HostResult,
  RangeFormat,
  RangeFormatData,
  TableInfo,
  TableUnlistInfo,
} from "./types";
import { unsupported } from "./types";

const EVIDENCE =
  "In-repo bridge only covers Application/ActiveWorkbook/JSIDE CodeModule; no verified format/table/chart contract";

/** Phase3 WPS capabilities without in-repo evidence → typed unsupported. */
export async function wpsReadFormat(
  _sheetName: string,
  _address: string,
): Promise<HostResult<RangeFormatData>> {
  return unsupported(
    "range.format.read",
    "wps-jsa",
    "Range format APIs are not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

export async function wpsWriteFormat(
  _sheetName: string,
  _address: string,
  _format: RangeFormat,
): Promise<HostResult<RangeFormatData>> {
  return unsupported(
    "range.format.write",
    "wps-jsa",
    "Range format APIs are not verified in this repository for WPS JSA",
    EVIDENCE,
  );
}

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

export async function wpsListConditionalFormats(
  _sheetName: string,
  _range: string,
): Promise<HostResult<never[]>> {
  return unsupported(
    "conditionalFormat.list",
    "wps-jsa",
    "Conditional formats are not verified for WPS JSA",
    "No in-repo WPS FormatConditions contract",
  ) as HostResult<never[]>;
}

export async function wpsAddConditionalFormat(_input: unknown): Promise<HostResult<never>> {
  return unsupported(
    "conditionalFormat.add",
    "wps-jsa",
    "Conditional formats are not verified for WPS JSA",
    "No in-repo WPS FormatConditions contract",
  ) as HostResult<never>;
}

export async function wpsDeleteConditionalFormat(
  _sheetName: string,
  _range: string,
  _id: string,
): Promise<HostResult<{ deleted: string }>> {
  return unsupported(
    "conditionalFormat.delete",
    "wps-jsa",
    "Conditional formats are not verified for WPS JSA",
    "No in-repo WPS FormatConditions contract",
  );
}

export async function wpsReadDataValidation(
  _sheetName: string,
  _range: string,
): Promise<HostResult<never>> {
  return unsupported(
    "dataValidation.read",
    "wps-jsa",
    "Data validation is not verified for WPS JSA",
    "No in-repo WPS Validation contract",
  ) as HostResult<never>;
}

export async function wpsWriteDataValidation(_input: unknown): Promise<HostResult<never>> {
  return unsupported(
    "dataValidation.write",
    "wps-jsa",
    "Data validation is not verified for WPS JSA",
    "No in-repo WPS Validation contract",
  ) as HostResult<never>;
}

export async function wpsClearDataValidation(
  _sheetName: string,
  _range: string,
): Promise<HostResult<{ cleared: string }>> {
  return unsupported(
    "dataValidation.clear",
    "wps-jsa",
    "Data validation is not verified for WPS JSA",
    "No in-repo WPS Validation contract",
  );
}

function wpsStructureUnsupported(capability: string): HostResult<never> {
  return unsupported(
    capability,
    "wps-jsa",
    `${capability} is not verified for WPS JSA`,
    "No in-repo WPS sheet visibility/protection/Names contract",
  ) as HostResult<never>;
}

export async function wpsUpdateTable(_input: unknown) {
  return unsupported(
    "table.update",
    "wps-jsa",
    "Table update is not verified for WPS JSA",
    "No in-repo WPS ListObject update contract",
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

export async function wpsGetSheetVisibility(_sheetName: string) {
  return wpsStructureUnsupported("sheet.visibility.get");
}
export async function wpsSetSheetVisibility(_sheetName: string, _visibility: string) {
  return wpsStructureUnsupported("sheet.visibility.set");
}
export async function wpsGetSheetProtection(_sheetName: string) {
  return wpsStructureUnsupported("sheet.protection.get");
}
export async function wpsProtectSheet(_sheetName: string, _password?: string) {
  return wpsStructureUnsupported("sheet.protection.protect");
}
export async function wpsUnprotectSheet(_sheetName: string, _password?: string) {
  return wpsStructureUnsupported("sheet.protection.unprotect");
}
export async function wpsListNamedRanges(_input?: unknown) {
  return wpsStructureUnsupported("namedRange.list");
}
export async function wpsCreateNamedRange(_input: unknown) {
  return wpsStructureUnsupported("namedRange.create");
}
export async function wpsUpdateNamedRange(_input: unknown) {
  return wpsStructureUnsupported("namedRange.update");
}
export async function wpsDeleteNamedRange(_input: unknown) {
  return wpsStructureUnsupported("namedRange.delete");
}

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
