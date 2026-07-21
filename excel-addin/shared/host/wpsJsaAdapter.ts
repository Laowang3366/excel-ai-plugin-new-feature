import { wpsInspectWorkbookObjects } from "./wpsJsaWorkbookObjects";
import { absoluteA1FromOrigin } from "./a1Address";
import { readWpsAddress } from "./wpsJsaAddress";
import { wpsGetSelection } from "./wpsJsaSelection";
import {
  formulaMatrixFrom,
  getApplication,
  getSheet,
  matrixFrom,
  requireWorkbook,
} from "./wpsJsaRuntime";
import {
  wpsCreateChart,
  wpsCreateTable,
  wpsDeleteChart,
  wpsDeleteTable,
  wpsListCharts,
  wpsListTables,
  wpsReadFormat,
  wpsUnlistTable,
  wpsWriteFormat,
} from "./wpsJsaUnsupported";
import {
  wpsAddConditionalFormat,
  wpsDeleteConditionalFormat,
  wpsListConditionalFormats,
} from "./wpsJsaConditionalFormat";
import {
  wpsClearDataValidation,
  wpsReadDataValidation,
  wpsWriteDataValidation,
} from "./wpsJsaDataValidation";
import { wpsInspectWorkbook } from "./wpsJsaInspect";
import { wpsSaveWorkbook } from "./wpsJsaWorkbookSave";
import {
  wpsCreatePivot,
  wpsListPivots,
  wpsRefreshPivots,
} from "./wpsJsaPivotUnsupported";
import { WpsJsaSlicerSupport } from "./wpsJsaSlicerUnsupported";
import { wpsCopySheet, wpsMoveSheet } from "./wpsJsaSheetOps";
import { wpsReadRange } from "./wpsJsaRangeRead";
import { wpsStructureSurface } from "./wpsJsaStructure";
import {
  wpsConvertFormulasToValues,
  wpsInspectFormulaDependencies,
  wpsRepairFormulaReferences,
} from "./wpsJsaFormulaGovernance";
import {
  wpsInspectFormulaBackups,
  wpsRestoreFormulas,
} from "./wpsJsaFormulaGovernanceBackups";
import type {
  CellValue,
  FormulaContextData,
  FormulaContextEntry,
  HostAdapter,
  HostResult,
  HostStatus,
  RangeData,
  SheetInfo,
} from "./types";
import { ok, unsupported } from "./types";
/** Core WPS JSA adapter: status/selection/range/formula/sheet + unsupported phase3 caps. */
export class WpsJsaAdapter extends WpsJsaSlicerSupport implements HostAdapter {
  readonly kind = "wps-jsa" as const;

  getRuntimeCapabilities() {
    return { dynamicArrayFunctionsEnabled: false };
  }

  async getStatus(): Promise<HostResult<HostStatus>> {
    const app = getApplication();
    if (!app) {
      return ok({
        kind: "wps-jsa",
        connected: false,
        hostName: "WPS 表格 (JSA)",
        workbookName: null,
        detail: "Application unavailable",
      });
    }
    return ok({
      kind: "wps-jsa",
      connected: Boolean(app.ActiveWorkbook),
      hostName: app.Name ?? "WPS 表格 (JSA)",
      workbookName: app.ActiveWorkbook?.Name ?? null,
    });
  }

  getSelection = wpsGetSelection;

  readRange = wpsReadRange;

  async getFormulaContext(
    sheetName: string,
    address?: string,
  ): Promise<HostResult<FormulaContextData>> {
    const workbookResult = requireWorkbook("formula.context");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet?.Range) {
      return unsupported(
        "formula.context",
        "wps-jsa",
        `Sheet "${sheetName}" or Range API missing`,
        "Assumed Range.Value2/Formula",
      );
    }
    let targetAddress = address?.trim() || "";
    if (!targetAddress) {
      const usedAddr = readWpsAddress(sheet.UsedRange);
      if (!usedAddr) {
        return unsupported(
          "formula.context",
          "wps-jsa",
          "range omitted and UsedRange unavailable",
          "Desktop uses UsedRange when address is empty",
        );
      }
      targetAddress = usedAddr;
    }
    const range = sheet.Range(targetAddress);
    const formulasMatrix = formulaMatrixFrom(range.Formula);
    const values = matrixFrom(range.Value2);
    const formulas: FormulaContextEntry[] = [];
    const origin = (readWpsAddress(range, targetAddress) ?? targetAddress).replace(/^.*!/, "");
    for (let r = 0; r < formulasMatrix.length; r += 1) {
      for (let c = 0; c < (formulasMatrix[r]?.length ?? 0); c += 1) {
        const formula = formulasMatrix[r][c] ?? "";
        if (!formula.startsWith("=")) continue;
        const cellAddress = absoluteA1FromOrigin(origin, r, c);
        formulas.push({
          address: cellAddress,
          formula,
          value: values[r]?.[c] ?? null,
        });
      }
    }
    return ok({
      sheetName,
      address: origin,
      formulas,
      cells: formulas,
    });
  }

  async writeRange(
    sheetName: string,
    address: string,
    values: CellValue[][],
  ): Promise<HostResult<RangeData>> {
    const workbookResult = requireWorkbook("range.write");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet?.Range) {
      return unsupported(
        "range.write",
        "wps-jsa",
        `Sheet "${sheetName}" or Range API missing`,
        "Assumed Worksheets.Item(name).Range(address).Value2 (not in bridge contract)",
      );
    }
    const range = sheet.Range(address);
    range.Value2 = values;
    return ok({
      sheetName,
      address: readWpsAddress(range, address) ?? address,
      values: matrixFrom(range.Value2),
      formulas: formulaMatrixFrom(range.Formula),
    });
  }

  async writeFormulas(
    sheetName: string,
    address: string,
    formulas: string[][],
  ): Promise<HostResult<RangeData>> {
    const workbookResult = requireWorkbook("formula.write");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet?.Range) {
      return unsupported(
        "formula.write",
        "wps-jsa",
        `Sheet "${sheetName}" or Range API missing`,
        "Assumed Worksheets.Item(name).Range(address).Formula (not in bridge contract)",
      );
    }
    const range = sheet.Range(address);
    const payload =
      formulas.length === 1 && formulas[0]?.length === 1 ? formulas[0][0] : formulas;
    range.Formula = payload;
    return ok({
      sheetName,
      address: readWpsAddress(range, address) ?? address,
      values: matrixFrom(range.Value2),
      formulas: formulaMatrixFrom(range.Formula),
    });
  }

  async clearRange(
    sheetName: string,
    address: string,
  ): Promise<HostResult<{ cleared: string }>> {
    const workbookResult = requireWorkbook("range.clear");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet?.Range) {
      return unsupported(
        "range.clear",
        "wps-jsa",
        `Sheet "${sheetName}" or Range API missing`,
        "Assumed Range.Clear (not in bridge contract)",
      );
    }
    const range = sheet.Range(address);
    if (typeof range.Clear !== "function") {
      return unsupported(
        "range.clear",
        "wps-jsa",
        "Range.Clear is not a function",
        "Assumed Range.Clear (not in bridge contract)",
      );
    }
    range.Clear();
    return ok({ cleared: readWpsAddress(range, address) ?? address });
  }

  async listSheets(): Promise<HostResult<SheetInfo[]>> {
    const workbookResult = requireWorkbook("sheet.list");
    if (!workbookResult.ok) return workbookResult;
    const workbook = workbookResult.data;
    const sheets = workbook.Worksheets;
    if (!sheets || typeof sheets.Count !== "number") {
      return unsupported(
        "sheet.list",
        "wps-jsa",
        "Worksheets collection unavailable",
        "Assumed Worksheets.Count/Item (not in bridge contract)",
      );
    }
    const activeName = workbook.ActiveSheet?.Name;
    const result: SheetInfo[] = [];
    for (let i = 1; i <= sheets.Count; i += 1) {
      const sheet = sheets.Item(i);
      result.push({
        name: sheet.Name,
        index: sheet.Index ?? i,
        isActive: sheet.Name === activeName,
      });
    }
    return ok(result);
  }

  async addSheet(sheetName: string): Promise<HostResult<SheetInfo>> {
    const workbookResult = requireWorkbook("sheet.add");
    if (!workbookResult.ok) return workbookResult;
    const sheets = workbookResult.data.Worksheets;
    if (typeof sheets.Add !== "function") {
      return unsupported(
        "sheet.add",
        "wps-jsa",
        "Worksheets.Add is unavailable",
        "Assumed Worksheets.Add (not in bridge contract)",
      );
    }
    const sheet = sheets.Add();
    sheet.Name = sheetName;
    return ok({ name: sheet.Name, index: sheet.Index ?? sheets.Count, isActive: false });
  }

  async renameSheet(sheetName: string, newName: string): Promise<HostResult<SheetInfo>> {
    const workbookResult = requireWorkbook("sheet.rename");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet) {
      return unsupported(
        "sheet.rename",
        "wps-jsa",
        `Sheet "${sheetName}" not found`,
        "Assumed Worksheets.Item(name).Name",
      );
    }
    sheet.Name = newName;
    return ok({ name: sheet.Name, index: sheet.Index ?? 0, isActive: false });
  }

  async deleteSheet(sheetName: string): Promise<HostResult<{ deleted: string }>> {
    const workbookResult = requireWorkbook("sheet.delete");
    if (!workbookResult.ok) return workbookResult;
    const sheet = getSheet(workbookResult.data, sheetName);
    if (!sheet || typeof sheet.Delete !== "function") {
      return unsupported(
        "sheet.delete",
        "wps-jsa",
        `Sheet "${sheetName}" missing or Delete unavailable`,
        "Assumed Worksheet.Delete (not in bridge contract)",
      );
    }
    sheet.Delete();
    return ok({ deleted: sheetName });
  }
  copySheet = wpsCopySheet;
  moveSheet = wpsMoveSheet;
  insertRange = wpsStructureSurface.insertRange;
  deleteRange = wpsStructureSurface.deleteRange;
  autofitRange = wpsStructureSurface.autofitRange;
  readFormat = wpsReadFormat;
  writeFormat = wpsWriteFormat;
  listTables = wpsListTables;
  createTable = wpsCreateTable;
  deleteTable = wpsDeleteTable;
  unlistTable = wpsUnlistTable;
  updateTable = wpsStructureSurface.updateTable;
  inspectFormulaProtection = wpsStructureSurface.inspectFormulaProtection;
  manageFormulaProtection = wpsStructureSurface.manageFormulaProtection;
  inspectFormulaDependencies = wpsInspectFormulaDependencies;
  repairFormulaReferences = wpsRepairFormulaReferences;
  convertFormulasToValues = wpsConvertFormulasToValues;
  inspectFormulaBackups = wpsInspectFormulaBackups;
  restoreFormulas = wpsRestoreFormulas;
  getTableFilter = wpsStructureSurface.getTableFilter;
  applyTableFilter = wpsStructureSurface.applyTableFilter;
  clearTableFilter = wpsStructureSurface.clearTableFilter;
  getTableSort = wpsStructureSurface.getTableSort;
  applyTableSort = wpsStructureSurface.applyTableSort;
  clearTableSort = wpsStructureSurface.clearTableSort;
  listCharts = wpsListCharts;
  createChart = wpsCreateChart;
  deleteChart = wpsDeleteChart;
  updateChart = wpsStructureSurface.updateChart;
  listChartSeries = wpsStructureSurface.listChartSeries;
  updateChartSeries = wpsStructureSurface.updateChartSeries;
  deleteChartSeries = wpsStructureSurface.deleteChartSeries;
  addChartSeries = wpsStructureSurface.addChartSeries;
  updateChartSeriesValues = wpsStructureSurface.updateChartSeriesValues;
  updateChartSeriesBubbleSizes = wpsStructureSurface.updateChartSeriesBubbleSizes;
  listChartSeriesTrendlines = wpsStructureSurface.listChartSeriesTrendlines;
  addChartSeriesTrendline = wpsStructureSurface.addChartSeriesTrendline;
  updateChartSeriesTrendline = wpsStructureSurface.updateChartSeriesTrendline;
  deleteChartSeriesTrendline = wpsStructureSurface.deleteChartSeriesTrendline;
  updateChartSeriesTrendlineFormat = wpsStructureSurface.updateChartSeriesTrendlineFormat;
  updateChartSeriesMarkers = wpsStructureSurface.updateChartSeriesMarkers;
  getChartImage = wpsStructureSurface.getChartImage;
  getRangeImage = wpsStructureSurface.getRangeImage;
  updateChartSource = wpsStructureSurface.updateChartSource;
  updateChartAxis = wpsStructureSurface.updateChartAxis;
  updateChartDataLabels = wpsStructureSurface.updateChartDataLabels;
  updateChartSeriesAxisGroup = wpsStructureSurface.updateChartSeriesAxisGroup;
  inspectWorkbook = () => wpsInspectWorkbook(() => this.listSheets());
  inspectWorkbookObjects = (input?: import("./workbookObjectsTypes").WorkbookObjectsInspectInput) =>
    wpsInspectWorkbookObjects(input ?? {}, {
      listSheets: () => this.listSheets(),
      listNamedRanges: (opts) => this.listNamedRanges(opts),
    });
  saveWorkbook = wpsSaveWorkbook;
  listConditionalFormats = wpsListConditionalFormats;
  addConditionalFormat = wpsAddConditionalFormat;
  deleteConditionalFormat = wpsDeleteConditionalFormat;
  readDataValidation = wpsReadDataValidation;
  writeDataValidation = wpsWriteDataValidation;
  clearDataValidation = wpsClearDataValidation;
  getSheetVisibility = wpsStructureSurface.getSheetVisibility;
  setSheetVisibility = wpsStructureSurface.setSheetVisibility;
  getSheetProtection = wpsStructureSurface.getSheetProtection;
  protectSheet = wpsStructureSurface.protectSheet;
  unprotectSheet = wpsStructureSurface.unprotectSheet;
  listNamedRanges = wpsStructureSurface.listNamedRanges;
  createNamedRange = wpsStructureSurface.createNamedRange;
  updateNamedRange = wpsStructureSurface.updateNamedRange;
  deleteNamedRange = wpsStructureSurface.deleteNamedRange;
  getSheetDisplay = wpsStructureSurface.getSheetDisplay;
  setSheetDisplay = wpsStructureSurface.setSheetDisplay;
  getSheetFreeze = wpsStructureSurface.getSheetFreeze;
  setSheetFreeze = wpsStructureSurface.setSheetFreeze;
  getSheetPageLayout = wpsStructureSurface.getSheetPageLayout;
  setSheetPageLayout = wpsStructureSurface.setSheetPageLayout;
  listShapes = wpsStructureSurface.listShapes;
  createShape = wpsStructureSurface.createShape;
  deleteShape = wpsStructureSurface.deleteShape;
  updateShape = wpsStructureSurface.updateShape;
  listPivots = wpsListPivots;
  createPivot = wpsCreatePivot;
  refreshPivots = wpsRefreshPivots;
}
