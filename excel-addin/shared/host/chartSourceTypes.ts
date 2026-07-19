import type { ChartSeriesInfo } from "./chartSeriesTypes";

/** Office.js Chart.setData seriesBy subset. */
export type ChartSeriesBy = "auto" | "rows" | "columns";

export const CHART_SERIES_BY: readonly ChartSeriesBy[] = ["auto", "rows", "columns"];

export function isChartSeriesBy(value: unknown): value is ChartSeriesBy {
  return typeof value === "string" && (CHART_SERIES_BY as readonly string[]).includes(value);
}

export interface ChartSourceUpdateInput {
  sheetName: string;
  chartName: string;
  sourceRange: string;
  /** Defaults to auto when omitted. */
  seriesBy?: ChartSeriesBy;
}

/** Real readback after setData + sync + series load. */
export interface ChartSourceInfo {
  sheetName: string;
  chartName: string;
  /** Same-sheet A1 range (sheet prefix stripped when present and matching). */
  sourceRange: string;
  seriesBy: ChartSeriesBy;
  series: ChartSeriesInfo[];
}
