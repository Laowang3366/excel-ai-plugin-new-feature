import { mapChartType, toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartSeriesInfo, ChartSeriesUpdateInput } from "./chartSeriesTypes";
import type { HostResult } from "./types";

/** Office.js ChartSeries surface used by this batch. */
interface ExcelChartSeries {
  name: string;
  chartType: string;
  smooth: boolean;
  load(props: string): void;
}

interface ExcelChartSeriesCollection {
  items: ExcelChartSeries[];
  getItemAt(index: number): ExcelChartSeries;
  load(props: string): void;
}

function chartSeries(chart: object): ExcelChartSeriesCollection {
  return (chart as { series: ExcelChartSeriesCollection }).series;
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") {
    throw new Error(`${field} is not a loaded string`);
  }
  return value;
}

function requireLoadedBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a loaded boolean`);
  }
  return value;
}

function toSeriesInfo(item: ExcelChartSeries, index: number): ChartSeriesInfo {
  return {
    index,
    name: requireLoadedString(item.name, "ChartSeries.name"),
    chartType: toChartTypeLabel(requireLoadedString(item.chartType, "ChartSeries.chartType")),
    smooth: requireLoadedBoolean(item.smooth, "ChartSeries.smooth"),
  };
}

/** List series for one chart; index is 1-based from real items after load+sync. */
export async function officeJsListChartSeries(
  sheetName: string,
  chartName: string,
): Promise<HostResult<ChartSeriesInfo[]>> {
  return withExcel("chart.series.list", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const series = chartSeries(sheet.charts.getItem(chartName));
    series.load("items/name,items/chartType,items/smooth");
    await context.sync();
    return series.items.map((item, i) => toSeriesInfo(item, i + 1));
  });
}

/** Update one series by 1-based index; write → sync → load+sync readback. */
export async function officeJsUpdateChartSeries(
  input: ChartSeriesUpdateInput,
): Promise<HostResult<ChartSeriesInfo>> {
  return withExcel("chart.series.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const collection = chartSeries(sheet.charts.getItem(input.chartName));
    const series = collection.getItemAt(input.seriesIndex - 1);
    if (input.newName != null) series.name = input.newName;
    if (input.chartType != null) series.chartType = mapChartType(input.chartType);
    if (input.smooth != null) series.smooth = input.smooth;
    await context.sync();
    series.load("name,chartType,smooth");
    await context.sync();
    return toSeriesInfo(series, input.seriesIndex);
  });
}
