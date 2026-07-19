import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartSeriesInfo } from "./chartSeriesTypes";
import type {
  ChartSeriesBy,
  ChartSourceInfo,
  ChartSourceUpdateInput,
} from "./chartSourceTypes";
import type { HostResult } from "./types";

interface ExcelChartSeries {
  name: string;
  chartType: string;
  smooth: boolean;
}

interface ExcelChartWithSetData {
  name: string;
  setData(range: object, seriesBy?: string): void;
  load(props: string): void;
  series: {
    items: ExcelChartSeries[];
    load(props: string): void;
  };
}

const SERIES_BY_OFFICE: Record<ChartSeriesBy, string> = {
  auto: "Auto",
  rows: "Rows",
  columns: "Columns",
};

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

/**
 * Same-sheet A1 only: bare range or matching Sheet!A1.
 * Rejects empty, cross-sheet, and non-A1-like addresses.
 */
export function normalizeSameSheetSourceRange(sheetName: string, sourceRange: string): string {
  const raw = sourceRange.trim();
  if (raw === "") throw new Error("sourceRange must be non-empty");
  let bare = raw;
  if (raw.includes("!")) {
    const bang = raw.lastIndexOf("!");
    const sheetPart = raw
      .slice(0, bang)
      .replace(/^'/, "")
      .replace(/'$/, "")
      .replace(/''/g, "'");
    bare = raw.slice(bang + 1).trim();
    if (sheetPart.toLowerCase() !== sheetName.toLowerCase()) {
      throw new Error("sourceRange must be on the same worksheet as the chart");
    }
  }
  bare = bare.replace(/\$/g, "").trim();
  if (bare === "") throw new Error("sourceRange must be non-empty");
  // Single cell or single contiguous range only (no multi-area comma lists).
  if (bare.includes(",")) {
    throw new Error("sourceRange multi-area is not supported");
  }
  if (!/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(bare)) {
    throw new Error("sourceRange must be a same-sheet A1 address");
  }
  for (const part of bare.split(":")) {
    const rowMatch = /^[A-Za-z]+(\d+)$/.exec(part);
    if (!rowMatch || Number(rowMatch[1]) < 1) {
      throw new Error("sourceRange row must be >= 1");
    }
  }
  return bare.toUpperCase();
}

function toSeriesInfo(item: ExcelChartSeries, index: number): ChartSeriesInfo {
  return {
    index,
    name: requireLoadedString(item.name, "ChartSeries.name"),
    chartType: toChartTypeLabel(requireLoadedString(item.chartType, "ChartSeries.chartType")),
    smooth: requireLoadedBoolean(item.smooth, "ChartSeries.smooth"),
  };
}

/** Replace chart data source via Chart.setData; return real series snapshot. */
export async function officeJsUpdateChartSource(
  input: ChartSourceUpdateInput,
): Promise<HostResult<ChartSourceInfo>> {
  return withExcel("chart.source.update", async (context) => {
    const seriesBy: ChartSeriesBy = input.seriesBy ?? "auto";
    const sourceRange = normalizeSameSheetSourceRange(input.sheetName, input.sourceRange);
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as ExcelChartWithSetData;
    const range = sheet.getRange(sourceRange);
    chart.setData(range, SERIES_BY_OFFICE[seriesBy]);
    await context.sync();
    chart.series.load("items/name,items/chartType,items/smooth");
    chart.load("name");
    await context.sync();
    const series = chart.series.items.map((item, i) => toSeriesInfo(item, i + 1));
    return {
      sheetName: input.sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      sourceRange,
      seriesBy,
      series,
    };
  });
}
