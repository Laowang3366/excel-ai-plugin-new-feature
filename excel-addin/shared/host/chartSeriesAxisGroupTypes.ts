import type { ChartAxisGroup } from "./chartAxisTypes";

export interface ChartSeriesAxisGroupUpdateInput {
  sheetName: string;
  chartName: string;
  /** Public 1-based series index. */
  seriesIndex: number;
  axisGroup: ChartAxisGroup;
}

/** Real snapshot after write→sync→load→sync. */
export interface ChartSeriesAxisGroupInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  /** Mapped primary|secondary or raw host string. */
  axisGroup: ChartAxisGroup | string;
}
