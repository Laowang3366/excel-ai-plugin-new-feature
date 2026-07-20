import type { ChartSeriesInfo } from "./chartSeriesTypes";

/** Office.js Chart.setData seriesBy subset. */
export type ChartSeriesBy = "auto" | "rows" | "columns";

export const CHART_SERIES_BY: readonly ChartSeriesBy[] = ["auto", "rows", "columns"];

export function isChartSeriesBy(value: unknown): value is ChartSeriesBy {
  return typeof value === "string" && (CHART_SERIES_BY as readonly string[]).includes(value);
}

export interface ChartSourceUpdateInput {
  /** Worksheet that owns the chart. */
  sheetName: string;
  chartName: string;
  /**
   * Data range for Chart.setData: bare A1 on chart sheet, or same-workbook
   * Sheet2!A1:B10 / 'Sheet 2'!A1:B10. External / 3D / multi-area / structured refs rejected.
   */
  sourceRange: string;
  /** Defaults to auto when omitted. */
  seriesBy?: ChartSeriesBy;
}

/** Real readback after setData + sync + series load. */
export interface ChartSourceInfo {
  sheetName: string;
  chartName: string;
  /**
   * Canonical source: bare A1 when source is on chart sheet;
   * Sheet!A1 or 'Sheet N'!A1 when cross-sheet.
   */
  sourceRange: string;
  seriesBy: ChartSeriesBy;
  series: ChartSeriesInfo[];
}
