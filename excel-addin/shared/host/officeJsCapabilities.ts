import { mapChartType, toChartTypeLabel } from "./officeJsChartTypes";
import {
  loadRangeFormat,
  readFormatFromRange,
  withExcel,
} from "./officeJsRuntime";
import type {
  ChartInfo,
  ChartType,
  HostResult,
  RangeFormat,
  RangeFormatData,
  TableInfo,
  WorkbookInspectInfo,
} from "./types";

/** Phase3 Office.js capabilities (format/table/chart/workbook). */
export async function officeJsReadFormat(
  sheetName: string,
  address: string,
): Promise<HostResult<RangeFormatData>> {
  return withExcel("range.format.read", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);
    loadRangeFormat(range);
    await context.sync();
    return {
      sheetName,
      address: range.address,
      format: readFormatFromRange(range),
    };
  });
}

export async function officeJsWriteFormat(
  sheetName: string,
  address: string,
  format: RangeFormat,
): Promise<HostResult<RangeFormatData>> {
  return withExcel("range.format.write", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);
    range.load("rowCount,columnCount");
    await context.sync();
    if (format.fontName != null) range.format.font.name = format.fontName;
    if (format.fontSize != null) range.format.font.size = format.fontSize;
    if (format.fontBold != null) range.format.font.bold = format.fontBold;
    if (format.fontColor != null) range.format.font.color = format.fontColor;
    if (format.fillColor != null) range.format.fill.color = format.fillColor;
    if (format.numberFormat != null) {
      const rows = Math.max(1, range.rowCount || 1);
      const cols = Math.max(1, range.columnCount || 1);
      range.numberFormat = Array.from({ length: rows }, () =>
        Array.from({ length: cols }, () => format.numberFormat as string),
      );
    }
    if (format.horizontalAlignment != null) {
      range.format.horizontalAlignment = format.horizontalAlignment;
    }
    if (format.verticalAlignment != null) {
      range.format.verticalAlignment = format.verticalAlignment;
    }
    if (format.wrapText != null) range.format.wrapText = format.wrapText;
    loadRangeFormat(range);
    await context.sync();
    return {
      sheetName,
      address: range.address,
      format: readFormatFromRange(range),
    };
  });
}

export async function officeJsListTables(
  sheetName?: string,
): Promise<HostResult<TableInfo[]>> {
  return withExcel("table.list", async (context) => {
    if (sheetName) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.tables.load("items/name,items/showHeaders,items/showFilterButton");
      await context.sync();
      const result: TableInfo[] = [];
      for (const table of sheet.tables.items) {
        const range = table.getRange();
        range.load("address");
        await context.sync();
        result.push({
          name: table.name,
          sheetName,
          address: range.address,
          hasHeaders: table.showHeaders,
          showFilter: table.showFilterButton,
        });
      }
      return result;
    }
    context.workbook.worksheets.load("items/name");
    await context.sync();
    const result: TableInfo[] = [];
    for (const sheet of context.workbook.worksheets.items) {
      sheet.tables.load("items/name,items/showHeaders,items/showFilterButton");
      await context.sync();
      for (const table of sheet.tables.items) {
        const range = table.getRange();
        range.load("address");
        await context.sync();
        result.push({
          name: table.name,
          sheetName: sheet.name,
          address: range.address,
          hasHeaders: table.showHeaders,
          showFilter: table.showFilterButton,
        });
      }
    }
    return result;
  });
}

export async function officeJsCreateTable(input: {
  sheetName: string;
  address: string;
  name?: string;
  hasHeaders?: boolean;
}): Promise<HostResult<TableInfo>> {
  return withExcel("table.create", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.add(input.address, input.hasHeaders !== false);
    if (input.name) table.name = input.name;
    table.load("name,showHeaders,showFilterButton");
    const range = table.getRange();
    range.load("address");
    await context.sync();
    return {
      name: table.name,
      sheetName: input.sheetName,
      address: range.address,
      hasHeaders: table.showHeaders,
      showFilter: table.showFilterButton,
    };
  });
}

export async function officeJsDeleteTable(
  sheetName: string,
  tableName: string,
): Promise<HostResult<{ deleted: string }>> {
  return withExcel("table.delete", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.tables.getItem(tableName).delete();
    await context.sync();
    return { deleted: tableName };
  });
}

export async function officeJsListCharts(
  sheetName?: string,
): Promise<HostResult<ChartInfo[]>> {
  return withExcel("chart.list", async (context) => {
    const result: ChartInfo[] = [];
    if (sheetName) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.load("name");
      sheet.charts.load(
        "items/name,items/chartType,items/style,items/left,items/top,items/width,items/height",
      );
      await context.sync();
      for (const chart of sheet.charts.items) {
        chart.title.load("text");
        chart.legend.load("visible");
        await context.sync();
        result.push({
          name: chart.name,
          sheetName: sheet.name,
          chartType: toChartTypeLabel(String(chart.chartType)),
          title: chart.title.text,
          left: chart.left,
          top: chart.top,
          width: chart.width,
          height: chart.height,
          style: chart.style,
          legendVisible: chart.legend.visible,
        });
      }
      return result;
    }
    context.workbook.worksheets.load("items/name");
    await context.sync();
    for (const sheet of context.workbook.worksheets.items) {
      sheet.charts.load(
        "items/name,items/chartType,items/style,items/left,items/top,items/width,items/height",
      );
      await context.sync();
      for (const chart of sheet.charts.items) {
        chart.title.load("text");
        chart.legend.load("visible");
        await context.sync();
        result.push({
          name: chart.name,
          sheetName: sheet.name,
          chartType: toChartTypeLabel(String(chart.chartType)),
          title: chart.title.text,
          left: chart.left,
          top: chart.top,
          width: chart.width,
          height: chart.height,
          style: chart.style,
          legendVisible: chart.legend.visible,
        });
      }
    }
    return result;
  });
}

export async function officeJsCreateChart(input: {
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
  return withExcel("chart.create", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const source = sheet.getRange(input.sourceRange);
    const chart = sheet.charts.add(mapChartType(input.chartType), source, "Auto");
    if (input.name) chart.name = input.name;
    if (input.title != null) {
      chart.title.text = input.title;
      chart.title.visible = true;
    }
    if (input.left != null) chart.left = input.left;
    if (input.top != null) chart.top = input.top;
    if (input.width != null) chart.width = input.width;
    if (input.height != null) chart.height = input.height;
    await context.sync();
    chart.load("name,chartType,style,left,top,width,height");
    chart.title.load("text,visible");
    chart.legend.load("visible");
    await context.sync();
    return {
      name: chart.name,
      sheetName: input.sheetName,
      chartType: toChartTypeLabel(String(chart.chartType)),
      title: chart.title.text,
      left: chart.left,
      top: chart.top,
      width: chart.width,
      height: chart.height,
      style: chart.style,
      legendVisible: chart.legend.visible,
    };
  });
}

export async function officeJsDeleteChart(
  sheetName: string,
  chartName: string,
): Promise<HostResult<{ deleted: string }>> {
  return withExcel("chart.delete", async (context) => {
    context.workbook.worksheets.getItem(sheetName).charts.getItem(chartName).delete();
    await context.sync();
    return { deleted: chartName };
  });
}

export async function officeJsInspectWorkbook(): Promise<HostResult<WorkbookInspectInfo>> {
  return withExcel("workbook.inspect", async (context) => {
    context.workbook.load("name");
    context.workbook.worksheets.load("items/name,items/position");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();

    const worksheets = context.workbook.worksheets.items;
    const usedBySheet = worksheets.map((sheet) => {
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("address,rowCount,columnCount");
      return used;
    });
    await context.sync();

    const sheets = worksheets.map((sheet, index) => {
      const used = usedBySheet[index]!;
      const empty = used.isNullObject;
      return {
        name: sheet.name,
        index: sheet.position,
        isActive: sheet.name === active.name,
        usedRangeAddress: empty ? null : used.address,
        rowCount: empty ? 0 : used.rowCount,
        columnCount: empty ? 0 : used.columnCount,
      };
    });
    const activeSheet = sheets.find((sheet) => sheet.isActive);
    return {
      workbookName: context.workbook.name,
      activeSheetName: active.name,
      sheetCount: sheets.length,
      usedRangeAddress: activeSheet?.usedRangeAddress ?? null,
      sheets,
    };
  });
}
