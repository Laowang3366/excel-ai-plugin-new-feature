import type {
  ChartAxisInfo,
  ChartAxisUpdateInput,
} from "./chartAxisTypes";
import type {
  ChartDataLabelsInfo,
  ChartDataLabelsUpdateInput,
} from "./chartDataLabelsTypes";
import type { ChartImageGetInput, ChartImageInfo } from "./chartImageTypes";
import type { RangeImageGetInput, RangeImageInfo } from "./rangeImageTypes";
import type {
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeInsertInput,
  RangeMutationInfo,
} from "./rangeStructureTypes";
import type {
  ChartSeriesAxisGroupInfo,
  ChartSeriesAxisGroupUpdateInput,
} from "./chartSeriesAxisGroupTypes";
import type {
  ChartSeriesBubbleSizesInfo,
  ChartSeriesBubbleSizesUpdateInput,
} from "./chartSeriesBubbleSizesTypes";
import type {
  ChartSeriesAddInput,
  ChartSeriesAddResult,
  ChartSeriesDeleteResult,
  ChartSeriesInfo,
  ChartSeriesUpdateInput,
} from "./chartSeriesTypes";
import type {
  ChartSeriesValuesInfo,
  ChartSeriesValuesUpdateInput,
} from "./chartSeriesValuesTypes";
import type { ChartSourceInfo, ChartSourceUpdateInput } from "./chartSourceTypes";
import type { ChartInfo, ChartType, ChartUpdateInput } from "./chartTypes";
import type { ShapeCreateInput, ShapeInfo, ShapeUpdateInput } from "./shapeTypes";
import type {
  WorkbookObjectsInspectInfo,
  WorkbookObjectsInspectInput,
} from "./workbookObjectsTypes";
import type { SheetDisplayInfo, SheetDisplayUpdateInput } from "./sheetDisplayTypes";
import type { SheetFreezeInfo, SheetFreezeSetInput } from "./sheetFreezeTypes";
import type {
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
} from "./sheetPageLayoutTypes";
import type {
  CellValue,
  ConditionalFormatInfo,
  ConditionalFormatRule,
  DataValidationInfo,
  DataValidationRule,
  FormulaContextData,
  HostKind,
  HostResult,
  HostRuntimeCapabilities,
  HostStatus,
  NamedRangeInfo,
  NamedRangeScope,
  RangeData,
  RangeExpandMode,
  RangeFormat,
  RangeFormatData,
  SelectionInfo,
  SheetInfo,
  SheetProtectionInfo,
  SheetVisibility,
  SheetVisibilityInfo,
  TableInfo,
  TableUnlistInfo,
  TableUpdateInput,
  WorkbookInspectInfo,
  WorkbookSaveInfo,
} from "./types";
import type {
  TableFilterApplyInput,
  TableFilterClearInput,
  TableFilterGetInput,
  TableFilterInfo,
} from "./tableFilterTypes";
import type {
  TableSortApplyInput,
  TableSortClearInput,
  TableSortGetInput,
  TableSortInfo,
} from "./tableSortTypes";
import type {
  FormulaProtectionInspectInfo,
  FormulaProtectionInspectInput,
  FormulaProtectionManageInfo,
  FormulaProtectionManageInput,
} from "./formulaProtectionTypes";
import type {
  FormulaBackupsInspectInfo,
  FormulaBackupsRestoreInfo,
  FormulaConvertToValuesInfo,
  FormulaConvertToValuesInput,
  FormulaDependenciesInspectInfo,
  FormulaDependenciesInspectInput,
  FormulaReferencesRepairInfo,
  FormulaReferencesRepairInput,
} from "./formulaGovernanceTypes";
import type {
  PivotCreateInfo,
  PivotCreateInput,
  PivotListInfo,
  PivotListInput,
  PivotRefreshInfo,
  PivotRefreshInput,
} from "./pivotTypes";

/** Host capability surface implemented by Office.js / WPS / Mock adapters. */
export interface HostAdapter {
  readonly kind: HostKind;
  getStatus(): Promise<HostResult<HostStatus>>;
  getRuntimeCapabilities(): HostRuntimeCapabilities;
  getSelection(): Promise<HostResult<SelectionInfo>>;
  readRange(
    sheetName: string,
    address: string,
    expand?: RangeExpandMode,
  ): Promise<HostResult<RangeData>>;
  writeRange(
    sheetName: string,
    address: string,
    values: CellValue[][],
  ): Promise<HostResult<RangeData>>;
  writeFormulas(
    sheetName: string,
    address: string,
    formulas: string[][],
  ): Promise<HostResult<RangeData>>;
  clearRange(sheetName: string, address: string): Promise<HostResult<{ cleared: string }>>;
  insertRange(input: RangeInsertInput): Promise<HostResult<RangeMutationInfo>>;
  deleteRange(input: RangeDeleteInput): Promise<HostResult<RangeMutationInfo>>;
  autofitRange(input: RangeAutofitInput): Promise<HostResult<RangeAutofitInfo>>;
  getFormulaContext(
    sheetName: string,
    address?: string,
  ): Promise<HostResult<FormulaContextData>>;
  listSheets(): Promise<HostResult<SheetInfo[]>>;
  addSheet(sheetName: string): Promise<HostResult<SheetInfo>>;
  renameSheet(sheetName: string, newName: string): Promise<HostResult<SheetInfo>>;
  deleteSheet(sheetName: string): Promise<HostResult<{ deleted: string }>>;
  copySheet(sheetName: string, newName?: string): Promise<HostResult<SheetInfo>>;
  moveSheet(sheetName: string, position: number): Promise<HostResult<SheetInfo>>;

  readFormat(sheetName: string, address: string): Promise<HostResult<RangeFormatData>>;
  writeFormat(
    sheetName: string,
    address: string,
    format: RangeFormat,
  ): Promise<HostResult<RangeFormatData>>;

  listTables(sheetName?: string): Promise<HostResult<TableInfo[]>>;
  createTable(input: {
    sheetName: string;
    address: string;
    name?: string;
    hasHeaders?: boolean;
  }): Promise<HostResult<TableInfo>>;
  deleteTable(sheetName: string, tableName: string): Promise<HostResult<{ deleted: string }>>;
  unlistTable(sheetName: string, tableName: string): Promise<HostResult<TableUnlistInfo>>;
  updateTable(input: TableUpdateInput): Promise<HostResult<TableInfo>>;
  getTableFilter(input: TableFilterGetInput): Promise<HostResult<TableFilterInfo>>;
  applyTableFilter(input: TableFilterApplyInput): Promise<HostResult<TableFilterInfo>>;
  clearTableFilter(input: TableFilterClearInput): Promise<HostResult<TableFilterInfo>>;
  getTableSort(input: TableSortGetInput): Promise<HostResult<TableSortInfo>>;
  applyTableSort(input: TableSortApplyInput): Promise<HostResult<TableSortInfo>>;
  clearTableSort(input: TableSortClearInput): Promise<HostResult<TableSortInfo>>;
  inspectFormulaProtection(
    input: FormulaProtectionInspectInput,
  ): Promise<HostResult<FormulaProtectionInspectInfo>>;
  manageFormulaProtection(
    input: FormulaProtectionManageInput,
  ): Promise<HostResult<FormulaProtectionManageInfo>>;
  inspectFormulaDependencies(
    input: FormulaDependenciesInspectInput,
  ): Promise<HostResult<FormulaDependenciesInspectInfo>>;
  repairFormulaReferences(
    input: FormulaReferencesRepairInput,
  ): Promise<HostResult<FormulaReferencesRepairInfo>>;
  convertFormulasToValues(
    input: FormulaConvertToValuesInput,
  ): Promise<HostResult<FormulaConvertToValuesInfo>>;
  inspectFormulaBackups(): Promise<HostResult<FormulaBackupsInspectInfo>>;
  restoreFormulas(input: {
    backupId: string;
    removeAfterRestore?: boolean;
  }): Promise<HostResult<FormulaBackupsRestoreInfo>>;

  listCharts(sheetName?: string): Promise<HostResult<ChartInfo[]>>;
  createChart(input: {
    sheetName: string;
    sourceRange: string;
    chartType?: ChartType;
    name?: string;
    title?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  }): Promise<HostResult<ChartInfo>>;
  deleteChart(sheetName: string, chartName: string): Promise<HostResult<{ deleted: string }>>;
  updateChart(input: ChartUpdateInput): Promise<HostResult<ChartInfo>>;
  listChartSeries(sheetName: string, chartName: string): Promise<HostResult<ChartSeriesInfo[]>>;
  updateChartSeries(input: ChartSeriesUpdateInput): Promise<HostResult<ChartSeriesInfo>>;
  deleteChartSeries(
    sheetName: string,
    chartName: string,
    seriesIndex: number,
  ): Promise<HostResult<ChartSeriesDeleteResult>>;
  addChartSeries(input: ChartSeriesAddInput): Promise<HostResult<ChartSeriesAddResult>>;
  updateChartSeriesValues(
    input: ChartSeriesValuesUpdateInput,
  ): Promise<HostResult<ChartSeriesValuesInfo>>;
  updateChartSeriesBubbleSizes(
    input: ChartSeriesBubbleSizesUpdateInput,
  ): Promise<HostResult<ChartSeriesBubbleSizesInfo>>;
  getChartImage(input: ChartImageGetInput): Promise<HostResult<ChartImageInfo>>;
  getRangeImage(input: RangeImageGetInput): Promise<HostResult<RangeImageInfo>>;
  updateChartSource(input: ChartSourceUpdateInput): Promise<HostResult<ChartSourceInfo>>;
  updateChartAxis(input: ChartAxisUpdateInput): Promise<HostResult<ChartAxisInfo>>;
  updateChartDataLabels(
    input: ChartDataLabelsUpdateInput,
  ): Promise<HostResult<ChartDataLabelsInfo>>;
  updateChartSeriesAxisGroup(
    input: ChartSeriesAxisGroupUpdateInput,
  ): Promise<HostResult<ChartSeriesAxisGroupInfo>>;
  inspectWorkbook(): Promise<HostResult<WorkbookInspectInfo>>;
  inspectWorkbookObjects(
    input?: WorkbookObjectsInspectInput,
  ): Promise<HostResult<WorkbookObjectsInspectInfo>>;
  /** Save the current workbook in place (no path/saveAs). */
  saveWorkbook(): Promise<HostResult<WorkbookSaveInfo>>;
  listConditionalFormats(
    sheetName: string,
    range: string,
  ): Promise<HostResult<ConditionalFormatInfo[]>>;
  addConditionalFormat(input: {
    sheetName: string;
    range: string;
    rule: ConditionalFormatRule;
  }): Promise<HostResult<ConditionalFormatInfo>>;
  deleteConditionalFormat(
    sheetName: string,
    range: string,
    id: string,
  ): Promise<HostResult<{ deleted: string }>>;

  readDataValidation(
    sheetName: string,
    range: string,
  ): Promise<HostResult<DataValidationInfo>>;
  writeDataValidation(input: {
    sheetName: string;
    range: string;
    rule: DataValidationRule;
  }): Promise<HostResult<DataValidationInfo>>;
  clearDataValidation(
    sheetName: string,
    range: string,
  ): Promise<HostResult<{ cleared: string }>>;

  getSheetVisibility(sheetName: string): Promise<HostResult<SheetVisibilityInfo>>;
  setSheetVisibility(
    sheetName: string,
    visibility: SheetVisibility,
  ): Promise<HostResult<SheetVisibilityInfo>>;

  getSheetProtection(sheetName: string): Promise<HostResult<SheetProtectionInfo>>;
  protectSheet(
    sheetName: string,
    password?: string,
  ): Promise<HostResult<SheetProtectionInfo>>;
  unprotectSheet(
    sheetName: string,
    password?: string,
  ): Promise<HostResult<SheetProtectionInfo>>;

  listNamedRanges(input?: {
    scope?: NamedRangeScope;
    sheetName?: string;
  }): Promise<HostResult<NamedRangeInfo[]>>;
  createNamedRange(input: {
    name: string;
    refersTo: string;
    scope: NamedRangeScope;
    sheetName?: string;
    visible?: boolean;
  }): Promise<HostResult<NamedRangeInfo>>;
  updateNamedRange(input: {
    name: string;
    scope: NamedRangeScope;
    sheetName?: string;
    newName?: string;
    refersTo?: string;
    visible?: boolean;
  }): Promise<HostResult<NamedRangeInfo>>;
  deleteNamedRange(input: {
    name: string;
    scope: NamedRangeScope;
    sheetName?: string;
  }): Promise<HostResult<{ deleted: string }>>;
  getSheetDisplay(sheetName: string): Promise<HostResult<SheetDisplayInfo>>;
  setSheetDisplay(input: SheetDisplayUpdateInput): Promise<HostResult<SheetDisplayInfo>>;
  getSheetFreeze(sheetName: string): Promise<HostResult<SheetFreezeInfo>>;
  setSheetFreeze(input: SheetFreezeSetInput): Promise<HostResult<SheetFreezeInfo>>;
  getSheetPageLayout(sheetName: string): Promise<HostResult<SheetPageLayoutInfo>>;
  setSheetPageLayout(input: SheetPageLayoutUpdateInput): Promise<HostResult<SheetPageLayoutInfo>>;

  listShapes(sheetName?: string): Promise<HostResult<ShapeInfo[]>>;
  createShape(input: ShapeCreateInput): Promise<HostResult<ShapeInfo>>;
  deleteShape(sheetName: string, shapeName: string): Promise<HostResult<{ deleted: string }>>;
  updateShape(input: ShapeUpdateInput): Promise<HostResult<ShapeInfo>>;

  listPivots(input?: PivotListInput): Promise<HostResult<PivotListInfo>>;
  createPivot(input: PivotCreateInput): Promise<HostResult<PivotCreateInfo>>;
  refreshPivots(input?: PivotRefreshInput): Promise<HostResult<PivotRefreshInfo>>;
}
