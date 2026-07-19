import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartSeriesDeleteResult, ChartSeriesInfo } from "./chartSeriesTypes";
import type { HostResult } from "./types";

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function requireLoadedBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") throw new Error(`${field} is not a loaded boolean`);
  return value;
}

/** Delete one series; delete→sync→load remaining→sync. */
export async function officeJsDeleteChartSeries(
  sheetName: string,
  chartName: string,
  seriesIndex: number,
): Promise<HostResult<ChartSeriesDeleteResult>> {
  return withExcel("chart.series.delete", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const chart = sheet.charts.getItem(chartName) as unknown as {
      name: string;
      series: {
        items: Array<{ name: string; chartType: string; smooth: boolean }>;
        getItemAt(index: number): { delete(): void };
        load(props: string): void;
      };
      load(props: string): void;
    };
    chart.series.getItemAt(seriesIndex - 1).delete();
    await context.sync();
    chart.load("name");
    chart.series.load("items/name,items/chartType,items/smooth");
    await context.sync();
    const remainingSeries: ChartSeriesInfo[] = chart.series.items.map((item, i) => ({
      index: i + 1,
      name: requireLoadedString(item.name, "ChartSeries.name"),
      chartType: toChartTypeLabel(requireLoadedString(item.chartType, "ChartSeries.chartType")),
      smooth: requireLoadedBoolean(item.smooth, "ChartSeries.smooth"),
    }));
    return {
      sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      deletedSeriesIndex: seriesIndex,
      remainingSeries,
    };
  });
}
