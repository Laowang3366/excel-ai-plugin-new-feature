export type HostKind = "office-js" | "wps-jsa" | "unknown";

export type CellValue = string | number | boolean | null;

export interface HostStatus {
  kind: HostKind;
  connected: boolean;
  hostName: string;
  workbookName: string | null;
  detail?: string;
}

export interface SelectionInfo {
  sheetName: string;
  address: string;
  values: CellValue[][];
  formulas: string[][];
}

export type RangeExpandMode = "none" | "spill" | "currentArray" | "currentRegion";

export interface RangeData {
  sheetName: string;
  address: string;
  values: CellValue[][];
  formulas: string[][];
  expanded?: boolean;
  expandMode?: RangeExpandMode;
}

/** Desktop-compatible formula cell entry (ExcelFormulaService.GetContext). */
export interface FormulaContextEntry {
  address: string;
  formula: string;
  value: CellValue;
}

/** Desktop shape: { sheetName, address, formulas[] }. `cells` kept as alias for compatibility. */
export interface FormulaContextData {
  sheetName: string;
  address: string;
  formulas: FormulaContextEntry[];
  /** @deprecated prefer formulas — kept for interim callers */
  cells?: FormulaContextEntry[];
}

export interface SheetInfo {
  name: string;
  index: number;
  isActive: boolean;
  usedRangeAddress?: string | null;
  rowCount?: number;
  columnCount?: number;
}

export type HorizontalAlignment = "general" | "left" | "center" | "right" | "fill" | "justify";
export type VerticalAlignment = "top" | "center" | "bottom" | "justify";

export interface RangeFormat {
  fontName?: string | null;
  fontSize?: number | null;
  fontBold?: boolean | null;
  fontColor?: string | null;
  fillColor?: string | null;
  numberFormat?: string | null;
  horizontalAlignment?: HorizontalAlignment | string | null;
  verticalAlignment?: VerticalAlignment | string | null;
  wrapText?: boolean | null;
}

export interface RangeFormatData {
  sheetName: string;
  address: string;
  format: RangeFormat;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  address: string;
  hasHeaders: boolean;
  showFilter?: boolean;
  showTotals?: boolean;
  style?: string;
}

export interface TableUpdateInput {
  sheetName: string;
  tableName: string;
  newName?: string;
  style?: string;
  showHeaders?: boolean;
  showTotals?: boolean;
  showFilterButton?: boolean;
}

/** Result of table.unlist (Table.convertToRange); data cells retained. */
export interface TableUnlistInfo {
  sheetName: string;
  tableName: string;
  address: string;
  unlisted: true;
}

export interface WorkbookInspectInfo {
  workbookName: string;
  activeSheetName: string;
  sheetCount: number;
  usedRangeAddress: string | null;
  sheets: SheetInfo[];
}

/** Office.js ConditionalFormatType subset we commit to. */
export type ConditionalFormatKind = "cellValue" | "custom";

export type CellValueOperator =
  | "greaterThan"
  | "lessThan"
  | "equalTo"
  | "between"
  | "notBetween";

export interface ConditionalFormatRule {
  kind: ConditionalFormatKind;
  /** cellValue: formula1 required; between also formula2 */
  operator?: CellValueOperator;
  formula1?: string;
  formula2?: string;
  /** custom expression formula */
  formula?: string;
  fillColor?: string;
  fontColor?: string;
}

export interface ConditionalFormatInfo {
  id: string;
  sheetName: string;
  range: string;
  kind: ConditionalFormatKind;
  summary: string;
}

export type DataValidationType = "list" | "wholeNumber";
export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "equalTo"
  | "greaterThan"
  | "lessThan";

export interface DataValidationRule {
  type: DataValidationType;
  /** wholeNumber: required */
  operator?: DataValidationOperator;
  /** list source string or wholeNumber formula1 */
  formula1?: string;
  formula2?: string;
  /** preferred for list; each item must be a non-empty string */
  listValues?: string[];
  allowBlank?: boolean;
}

export interface DataValidationInfo {
  sheetName: string;
  range: string;
  rule: DataValidationRule | null;
}

/** Office.js Excel.SheetVisibility subset. */
export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export interface SheetVisibilityInfo {
  sheetName: string;
  visibility: SheetVisibility;
}

export interface SheetProtectionInfo {
  sheetName: string;
  protected: boolean;
}

export type { ChartInfo, ChartType, ChartUpdateInput } from "./chartTypes";
export { CHART_TYPES, isChartType } from "./chartTypes";
export type {
  ChartSeriesAddInput, ChartSeriesAddResult, ChartSeriesDeleteResult, ChartSeriesInfo, ChartSeriesUpdateInput,
} from "./chartSeriesTypes";
export type { ChartSeriesValuesInfo, ChartSeriesValuesUpdateInput } from "./chartSeriesValuesTypes";
export type {
  ChartSeriesBubbleSizesInfo,
  ChartSeriesBubbleSizesUpdateInput,
} from "./chartSeriesBubbleSizesTypes";
export type { ChartImageGetInput, ChartImageInfo } from "./chartImageTypes";
export type {
  GeometricShapeType, ShapeCreateInput, ShapeCreateKind, ShapeInfo, ShapeUpdateInput,
} from "./shapeTypes";
export { GEOMETRIC_SHAPE_TYPES, isGeometricShapeType } from "./shapeTypes";
export type { SheetDisplayInfo, SheetDisplayUpdateInput } from "./sheetDisplayTypes";
export type { SheetFreezeCommand, SheetFreezeInfo, SheetFreezeSetInput } from "./sheetFreezeTypes";
export type {
  PageOrientation,
  PagePaperSize,
  PageLayoutMargins,
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
} from "./sheetPageLayoutTypes";
export type { HostAdapter } from "./hostAdapter";

export type NamedRangeScope = "workbook" | "worksheet";

export interface NamedRangeInfo {
  name: string;
  refersTo: string;
  scope: NamedRangeScope;
  sheetName?: string;
  visible?: boolean;
}

export type SheetOperationKind = "add" | "rename" | "delete" | "copy" | "move";

export type SheetOperation =
  | { operation: "list" }
  | { operation: "add"; sheetName: string }
  | { operation: "rename"; sheetName: string; newName: string }
  | { operation: "delete"; sheetName: string }
  | { operation: "copy"; sheetName: string; newName?: string }
  | { operation: "move"; sheetName: string; position: number };

export interface UnsupportedResult {
  ok: false;
  unsupported: true;
  capability: string;
  host: HostKind;
  reason: string;
  evidence?: string;
}

/** Ordinary host/runtime failure (not typed capability unsupported). */
export interface FailResult {
  ok: false;
  unsupported?: false;
  reason: string;
  /** Absent on ordinary failures; kept optional so HostResult unions stay narrowable. */
  capability?: undefined;
  host?: undefined;
  evidence?: undefined;
}

export interface OkResult<T> {
  ok: true;
  data: T;
}

export type HostResult<T> = OkResult<T> | UnsupportedResult | FailResult;

export function unsupported(
  capability: string, host: HostKind, reason: string, evidence?: string,
): UnsupportedResult {
  return { ok: false, unsupported: true, capability, host, reason, evidence };
}
export function fail(reason: string): FailResult {
  return { ok: false, reason };
}
export function ok<T>(data: T): OkResult<T> { return { ok: true, data }; }
