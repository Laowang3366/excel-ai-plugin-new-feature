import type { ChartType } from "./chartTypes";

/** Public chart series snapshot (index is 1-based). */
export interface ChartSeriesInfo {
  index: number;
  name: string;
  chartType: ChartType | string;
  smooth: boolean;
}

/** Update one series; seriesIndex is public 1-based. */
export interface ChartSeriesUpdateInput {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  newName?: string;
  chartType?: ChartType;
  smooth?: boolean;
}

/** Result of chart.series.delete after real collection readback. */
export interface ChartSeriesDeleteResult {
  sheetName: string;
  chartName: string;
  deletedSeriesIndex: number;
  remainingSeries: ChartSeriesInfo[];
}

/** Create empty series (no values/xValues); dataBound always false. */
export interface ChartSeriesAddInput {
  sheetName: string;
  chartName: string;
  name?: string;
}

/** Result of chart.series.add after real collection readback. */
export interface ChartSeriesAddResult {
  sheetName: string;
  chartName: string;
  addedSeries: ChartSeriesInfo;
  dataBound: false;
}
