import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartSeriesAddInput, ChartSeriesAddResult, ChartSeriesInfo } from "./chartSeriesTypes";
import type { HostResult } from "./types";

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function requireLoadedBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") throw new Error(`${field} is not a loaded boolean`);
  return value;
}

/** Append empty series; load collection → add → sync → load last item → sync. */
export async function officeJsAddChartSeries(
  input: ChartSeriesAddInput,
): Promise<HostResult<ChartSeriesAddResult>> {
  return withExcel("chart.series.add", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      series: {
        items: Array<{ name: string; chartType: string; smooth: boolean }>;
        add(name?: string): unknown;
        load(props: string): void;
      };
      load(props: string): void;
    };
    chart.series.load("items");
    await context.sync();
    if (input.name != null) chart.series.add(input.name);
    else chart.series.add();
    await context.sync();
    chart.load("name");
    chart.series.load("items/name,items/chartType,items/smooth");
    await context.sync();
    const items = chart.series.items;
    if (items.length < 1) throw new Error("chart.series.add produced empty collection");
    const last = items[items.length - 1]!;
    const addedSeries: ChartSeriesInfo = {
      index: items.length,
      name: requireLoadedString(last.name, "ChartSeries.name"),
      chartType: toChartTypeLabel(requireLoadedString(last.chartType, "ChartSeries.chartType")),
      smooth: requireLoadedBoolean(last.smooth, "ChartSeries.smooth"),
    };
    return {
      sheetName: input.sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      addedSeries,
      dataBound: false,
    };
  });
}
