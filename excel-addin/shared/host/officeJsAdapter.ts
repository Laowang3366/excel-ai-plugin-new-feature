import {
  getExcelRun,
  normalizeFormulas,
  normalizeMatrix,
  withExcel,
} from "./officeJsRuntime";
import {
  officeJsCreateChart,
  officeJsCreateTable,
  officeJsDeleteChart,
  officeJsDeleteTable,
  officeJsInspectWorkbook,
  officeJsListCharts,
  officeJsListTables,
  officeJsReadFormat,
  officeJsWriteFormat,
} from "./officeJsCapabilities";
import { officeJsUnlistTable } from "./officeJsTableUnlist";
import {
  officeJsGetFormulaContext,
  officeJsReadRange,
} from "./officeJsRangeExpand";
import { officeJsCopySheet, officeJsMoveSheet } from "./officeJsSheetOps";
import {
  officeJsAddConditionalFormat,
  officeJsClearDataValidation,
  officeJsDeleteConditionalFormat,
  officeJsListConditionalFormats,
  officeJsReadDataValidation,
  officeJsWriteDataValidation,
} from "./officeJsValidation";
import {
  officeJsListChartSeries,
  officeJsUpdateChartSeries,
} from "./officeJsChartSeries";
import { officeJsAddChartSeries } from "./officeJsChartSeriesAdd";
import { officeJsDeleteChartSeries } from "./officeJsChartSeriesDelete";
import { officeJsUpdateChartSeriesValues } from "./officeJsChartSeriesValues";
import { officeJsUpdateChartSeriesBubbleSizes } from "./officeJsChartSeriesBubbleSizes";
import { officeJsGetChartImage } from "./officeJsChartImage";
import { officeJsGetRangeImage } from "./officeJsRangeImage";
import {
  officeJsAutofitRange,
  officeJsDeleteRange,
  officeJsInsertRange,
} from "./officeJsRangeStructure";
import { officeJsUpdateChartAxis } from "./officeJsChartAxes";
import { officeJsUpdateChartDataLabels } from "./officeJsChartDataLabels";
import { officeJsUpdateChartSeriesAxisGroup } from "./officeJsChartSeriesAxisGroup";
import { officeJsUpdateChartSource } from "./officeJsChartSource";
import {
  officeJsUpdateChart,
  officeJsUpdateTable,
} from "./officeJsObjectUpdate";
import {
  officeJsGetSheetFreeze,
  officeJsSetSheetFreeze,
} from "./officeJsFreeze";
import {
  officeJsGetSheetPageLayout,
  officeJsSetSheetPageLayout,
} from "./officeJsPageLayout";
import {
  officeJsCreateShape,
  officeJsDeleteShape,
  officeJsListShapes,
  officeJsUpdateShape,
} from "./officeJsShapes";
import {
  officeJsGetSheetDisplay,
  officeJsSetSheetDisplay,
} from "./officeJsSheetDisplay";
import {
  officeJsCreateNamedRange,
  officeJsDeleteNamedRange,
  officeJsGetSheetProtection,
  officeJsGetSheetVisibility,
  officeJsListNamedRanges,
  officeJsProtectSheet,
  officeJsSetSheetVisibility,
  officeJsUnprotectSheet,
  officeJsUpdateNamedRange,
} from "./officeJsStructure";
import type {
  CellValue,
  HostAdapter,
  HostResult,
  HostStatus,
  RangeData,
  SelectionInfo,
  SheetInfo,
} from "./types";
import { ok } from "./types";

/** Core Office.js adapter: status/selection/range/formula/sheet + phase3 capabilities. */
export class OfficeJsAdapter implements HostAdapter {
  readonly kind = "office-js" as const;

  async getStatus(): Promise<HostResult<HostStatus>> {
    const run = getExcelRun();
    if (!run) {
      return ok({
        kind: "office-js",
        connected: false,
        hostName: "Microsoft Excel (Office.js)",
        workbookName: null,
        detail: "Excel.run unavailable",
      });
    }
    return withExcel("host.status", async (context) => {
      context.workbook.load("name");
      await context.sync();
      return {
        kind: "office-js" as const,
        connected: true,
        hostName: "Microsoft Excel (Office.js)",
        workbookName: context.workbook.name,
      };
    });
  }

  async getSelection(): Promise<HostResult<SelectionInfo>> {
    return withExcel("selection.get", async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address,values,formulas");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      return {
        sheetName: sheet.name,
        address: range.address,
        values: normalizeMatrix(range.values),
        formulas: normalizeFormulas(range.formulas),
      };
    });
  }

  readRange = officeJsReadRange;
  getFormulaContext = officeJsGetFormulaContext;

  async writeRange(
    sheetName: string,
    address: string,
    values: CellValue[][],
  ): Promise<HostResult<RangeData>> {
    return withExcel("range.write", async (context) => {
      const range = context.workbook.worksheets.getItem(sheetName).getRange(address);
      range.values = values;
      range.load("address,values,formulas");
      await context.sync();
      return {
        sheetName,
        address: range.address,
        values: normalizeMatrix(range.values),
        formulas: normalizeFormulas(range.formulas),
      };
    });
  }

  async writeFormulas(
    sheetName: string,
    address: string,
    formulas: string[][],
  ): Promise<HostResult<RangeData>> {
    return withExcel("formula.write", async (context) => {
      const range = context.workbook.worksheets.getItem(sheetName).getRange(address);
      range.formulas = formulas;
      range.load("address,values,formulas");
      await context.sync();
      return {
        sheetName,
        address: range.address,
        values: normalizeMatrix(range.values),
        formulas: normalizeFormulas(range.formulas),
      };
    });
  }

  async clearRange(
    sheetName: string,
    address: string,
  ): Promise<HostResult<{ cleared: string }>> {
    return withExcel("range.clear", async (context) => {
      const range = context.workbook.worksheets.getItem(sheetName).getRange(address);
      range.clear();
      range.load("address");
      await context.sync();
      return { cleared: range.address };
    });
  }

  async listSheets(): Promise<HostResult<SheetInfo[]>> {
    return withExcel("sheet.list", async (context) => {
      context.workbook.worksheets.load("items/name,items/position");
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();
      return context.workbook.worksheets.items.map((sheet) => ({
        name: sheet.name,
        index: sheet.position,
        isActive: sheet.name === active.name,
      }));
    });
  }

  async addSheet(sheetName: string): Promise<HostResult<SheetInfo>> {
    return withExcel("sheet.add", async (context) => {
      const sheet = context.workbook.worksheets.add(sheetName);
      sheet.load("name,position");
      await context.sync();
      return { name: sheet.name, index: sheet.position, isActive: false };
    });
  }

  async renameSheet(sheetName: string, newName: string): Promise<HostResult<SheetInfo>> {
    return withExcel("sheet.rename", async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.name = newName;
      sheet.load("name,position");
      await context.sync();
      return { name: sheet.name, index: sheet.position, isActive: false };
    });
  }

  async deleteSheet(sheetName: string): Promise<HostResult<{ deleted: string }>> {
    return withExcel("sheet.delete", async (context) => {
      context.workbook.worksheets.getItem(sheetName).delete();
      await context.sync();
      return { deleted: sheetName };
    });
  }

  copySheet = officeJsCopySheet;
  moveSheet = officeJsMoveSheet;
  insertRange = officeJsInsertRange;
  deleteRange = officeJsDeleteRange;
  autofitRange = officeJsAutofitRange;

  readFormat = officeJsReadFormat;
  writeFormat = officeJsWriteFormat;
  listTables = officeJsListTables;
  createTable = officeJsCreateTable;
  deleteTable = officeJsDeleteTable;
  unlistTable = officeJsUnlistTable;
  updateTable = officeJsUpdateTable;
  listCharts = officeJsListCharts;
  createChart = officeJsCreateChart;
  deleteChart = officeJsDeleteChart;
  updateChart = officeJsUpdateChart;
  listChartSeries = officeJsListChartSeries;
  updateChartSeries = officeJsUpdateChartSeries;
  deleteChartSeries = officeJsDeleteChartSeries;
  addChartSeries = officeJsAddChartSeries;
  updateChartSeriesValues = officeJsUpdateChartSeriesValues;
  updateChartSeriesBubbleSizes = officeJsUpdateChartSeriesBubbleSizes;
  getChartImage = officeJsGetChartImage;
  getRangeImage = officeJsGetRangeImage;
  updateChartSource = officeJsUpdateChartSource;
  updateChartAxis = officeJsUpdateChartAxis;
  updateChartDataLabels = officeJsUpdateChartDataLabels;
  updateChartSeriesAxisGroup = officeJsUpdateChartSeriesAxisGroup;
  inspectWorkbook = officeJsInspectWorkbook;

  listConditionalFormats = officeJsListConditionalFormats;
  addConditionalFormat = officeJsAddConditionalFormat;
  deleteConditionalFormat = officeJsDeleteConditionalFormat;
  readDataValidation = officeJsReadDataValidation;
  writeDataValidation = officeJsWriteDataValidation;
  clearDataValidation = officeJsClearDataValidation;

  getSheetVisibility = officeJsGetSheetVisibility;
  setSheetVisibility = officeJsSetSheetVisibility;
  getSheetProtection = officeJsGetSheetProtection;
  protectSheet = officeJsProtectSheet;
  unprotectSheet = officeJsUnprotectSheet;
  listNamedRanges = officeJsListNamedRanges;
  createNamedRange = officeJsCreateNamedRange;
  updateNamedRange = officeJsUpdateNamedRange;
  deleteNamedRange = officeJsDeleteNamedRange;
  getSheetDisplay = officeJsGetSheetDisplay;
  setSheetDisplay = officeJsSetSheetDisplay;
  getSheetFreeze = officeJsGetSheetFreeze;
  setSheetFreeze = officeJsSetSheetFreeze;
  getSheetPageLayout = officeJsGetSheetPageLayout;
  setSheetPageLayout = officeJsSetSheetPageLayout;
  listShapes = officeJsListShapes;
  createShape = officeJsCreateShape;
  deleteShape = officeJsDeleteShape;
  updateShape = officeJsUpdateShape;
}
