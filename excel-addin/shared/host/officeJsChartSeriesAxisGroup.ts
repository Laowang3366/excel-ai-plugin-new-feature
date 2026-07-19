import { withExcel } from "./officeJsRuntime";
import type { ChartAxisGroup } from "./chartAxisTypes";
import type {
  ChartSeriesAxisGroupInfo,
  ChartSeriesAxisGroupUpdateInput,
} from "./chartSeriesAxisGroupTypes";
import type { HostResult } from "./types";

const GROUP_OFFICE: Record<ChartAxisGroup, string> = {
  primary: "Primary",
  secondary: "Secondary",
};

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

/** Map host axisGroup; known Primary/Secondary → public; unknown string passthrough. */
export function mapAxisGroupLabel(raw: unknown): ChartAxisGroup | string {
  if (typeof raw !== "string") {
    throw new Error("ChartSeries.axisGroup is not a loaded string");
  }
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "primary") return "primary";
  if (key === "secondary") return "secondary";
  return raw;
}

/** Update series.axisGroup; write → sync → load+sync real snapshot. */
export async function officeJsUpdateChartSeriesAxisGroup(
  input: ChartSeriesAxisGroupUpdateInput,
): Promise<HostResult<ChartSeriesAxisGroupInfo>> {
  return withExcel("chart.series.axisGroup.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      series: {
        getItemAt(index: number): { axisGroup: string; load(props: string): void };
      };
      load(props: string): void;
    };
    const series = chart.series.getItemAt(input.seriesIndex - 1);
    series.axisGroup = GROUP_OFFICE[input.axisGroup];
    await context.sync();
    chart.load("name");
    series.load("axisGroup");
    await context.sync();
    return {
      sheetName: input.sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      seriesIndex: input.seriesIndex,
      axisGroup: mapAxisGroupLabel(series.axisGroup),
    };
  });
}
