export type HostKind = "office-js" | "wps-jsa" | "unknown";

export type CellValue = string | number | boolean | null;

export interface HostStatus {
  kind: HostKind;
  connected: boolean;
  hostName: string;
  workbookName: string | null;
  detail?: string;
}

export interface HostRuntimeCapabilities {
  dynamicArrayFunctionsEnabled: boolean;
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
  showBandedRows?: boolean;
  showBandedColumns?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
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
  showBandedRows?: boolean;
  showBandedColumns?: boolean;
  /** Highlight first column (Excel.Table.showFirstColumn, ExcelApi 1.3). */
  showFirstColumn?: boolean;
  /** Highlight last column (Excel.Table.showLastColumn, ExcelApi 1.3). */
  showLastColumn?: boolean;
  /** Same-sheet single-area A1 range used by Excel.Table.resize. */
  resizeAddress?: string;
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

/** Result of workbook.save (current workbook only; no path/saveAs). */
export interface WorkbookSaveInfo {
  workbookName: string;
  saved: true;
}

/**
 * Writable CF kinds (add). List may also return kind "unsupported" for host types
 * that are recognized but not add-capable (DataBar, ColorScale, …).
 */
export type ConditionalFormatKind = "cellValue" | "custom";

/** List result kind includes unsupported host types (never silent cellValue). */
export type ConditionalFormatListKind = ConditionalFormatKind | "unsupported";

/** Public cellValue operators (complete Office.js ConditionalCellValueRule set). */
export type CellValueOperator =
  | "greaterThan"
  | "greaterThanOrEqualTo"
  | "lessThan"
  | "lessThanOrEqualTo"
  | "equalTo"
  | "notEqualTo"
  | "between"
  | "notBetween";

export interface ConditionalFormatRule {
  kind: ConditionalFormatKind;
  /** cellValue: operator + formula1; between/notBetween also formula2 */
  operator?: CellValueOperator;
  formula1?: string;
  formula2?: string;
  /** custom expression formula */
  formula?: string;
  /** #RRGGBB only */
  fillColor?: string;
  fontColor?: string;
}

export interface ConditionalFormatInfo {
  id: string;
  sheetName: string;
  range: string;
  /** Normalized: cellValue | custom | unsupported */
  kind: ConditionalFormatListKind;
  /** Host ConditionalFormatType string (CellValue, ContainsText, …). */
  hostType: string;
  /** true only for cellValue/custom that this add-in can add/verify. */
  supported: boolean;
  summary: string;
  limitations?: string[];
}

/**
 * Writable DV types aligned with Office.js DataValidationType (except None /
 * Inconsistent / MixedCriteria which are read-only host states).
 */
export type DataValidationType =
  | "list"
  | "wholeNumber"
  | "decimal"
  | "date"
  | "time"
  | "textLength"
  | "custom";

/** Complete Office.js DataValidationOperator set. */
export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "equalTo"
  | "notEqualTo"
  | "greaterThan"
  | "greaterThanOrEqualTo"
  | "lessThan"
  | "lessThanOrEqualTo";

export interface DataValidationRule {
  type: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  /** Inline list items only; mutually exclusive with formula1 range source. */
  listValues?: string[];
  allowBlank?: boolean;
}

export type DataValidationListSourceKind = "inline" | "range";

export interface DataValidationInfo {
  sheetName: string;
  range: string;
  rule: DataValidationRule | null;
  /** Host DataValidationType (List, WholeNumber, Inconsistent, …). */
  hostType?: string | null;
  /** false when host type is not a single writable rule (Inconsistent/MixedCriteria/unknown). */
  supported?: boolean;
  listSourceKind?: DataValidationListSourceKind | null;
  limitations?: string[];
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
export type {
  ChartMarkerStyle,
  ChartSeriesMarkersInfo,
  ChartSeriesMarkersUpdateInput,
} from "./chartSeriesMarkersTypes";
export { CHART_MARKER_STYLES, isChartMarkerStyle } from "./chartSeriesMarkersTypes";
export type { ChartImageGetInput, ChartImageInfo } from "./chartImageTypes";
export type { RangeImageGetInput, RangeImageInfo } from "./rangeImageTypes";
export type {
  GeometricShapeType, ShapeCreateInput, ShapeCreateKind, ShapeInfo, ShapeUpdateInput,
} from "./shapeTypes";
export { GEOMETRIC_SHAPE_TYPES, isGeometricShapeType } from "./shapeTypes";
export type { SheetDisplayInfo, SheetDisplayUpdateInput } from "./sheetDisplayTypes";
export type { SheetFreezeCommand, SheetFreezeInfo, SheetFreezeSetInput } from "./sheetFreezeTypes";
export type {
  RangeAutofitDirection,
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeDeleteShift,
  RangeInsertInput,
  RangeInsertShift,
  RangeMutationInfo,
} from "./rangeStructureTypes";
export type {
  PageOrientation,
  PagePaperSize,
  PageOrder,
  PageLayoutMargins,
  PageLayoutTextSides,
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
} from "./sheetPageLayoutTypes";
export type { HostAdapter } from "./hostAdapter";
export type {
  PivotAggregationFunction,
  PivotCreateInfo,
  PivotCreateInput,
  PivotListInfo,
  PivotListInput,
  PivotRefreshInfo,
  PivotRefreshInput,
  PivotTableInfo,
} from "./pivotTypes";

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
  capability: string;
  host: HostKind;
  reason: string;
  evidence?: string;
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
export function fail(
  capability: string, host: HostKind, reason: string, evidence?: string,
): FailResult {
  return { ok: false, unsupported: false, capability, host, reason, evidence };
}
export function ok<T>(data: T): OkResult<T> { return { ok: true, data }; }
