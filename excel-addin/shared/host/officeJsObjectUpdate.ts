import { mapChartType, toChartTypeLabel } from "./officeJsChartTypes";
import { normalizeSameSheetSourceRange } from "./officeJsChartSource";
import { withExcel } from "./officeJsRuntime";
import type {
  ChartInfo,
  ChartUpdateInput,
  HostResult,
  TableInfo,
  TableUpdateInput,
} from "./types";

export async function officeJsUpdateTable(
  input: TableUpdateInput,
): Promise<HostResult<TableInfo>> {
  return withExcel("table.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    if (input.resizeAddress != null) {
      table.resize(normalizeSameSheetSourceRange(input.sheetName, input.resizeAddress));
    }
    if (input.newName != null) table.name = input.newName;
    if (input.style != null) table.style = input.style;
    if (input.showHeaders != null) table.showHeaders = input.showHeaders;
    if (input.showTotals != null) table.showTotals = input.showTotals;
    if (input.showFilterButton != null) table.showFilterButton = input.showFilterButton;
    if (input.showBandedRows != null) table.showBandedRows = input.showBandedRows;
    if (input.showBandedColumns != null) table.showBandedColumns = input.showBandedColumns;
    await context.sync();
    table.load(
      "name,showHeaders,showFilterButton,showTotals,showBandedRows,showBandedColumns,style",
    );
    const range = table.getRange();
    range.load("address");
    await context.sync();
    return {
      name: table.name,
      sheetName: input.sheetName,
      address: range.address,
      hasHeaders: table.showHeaders,
      showFilter: table.showFilterButton,
      showTotals: table.showTotals,
      showBandedRows: table.showBandedRows,
      showBandedColumns: table.showBandedColumns,
      style: table.style,
    };
  });
}

export async function officeJsUpdateChart(
  input: ChartUpdateInput,
): Promise<HostResult<ChartInfo>> {
  return withExcel("chart.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName);
    if (input.newName != null) chart.name = input.newName;
    if (input.chartType != null) chart.chartType = mapChartType(input.chartType);
    if (input.title != null) {
      chart.title.text = input.title;
      chart.title.visible = input.showTitle !== false;
    } else if (input.showTitle != null) {
      chart.title.visible = input.showTitle;
    }
    if (input.left != null) chart.left = input.left;
    if (input.top != null) chart.top = input.top;
    if (input.width != null) chart.width = input.width;
    if (input.height != null) chart.height = input.height;
    if (input.style != null) chart.style = input.style;
    if (input.showLegend != null) chart.legend.visible = input.showLegend;
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
      titleVisible: chart.title.visible,
      left: chart.left,
      top: chart.top,
      width: chart.width,
      height: chart.height,
      style: chart.style,
      legendVisible: chart.legend.visible,
    };
  });
}
