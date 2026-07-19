/** Office.js chart types (Phase13 basics + Phase27 deep). */
export type ChartType =
  | "column"
  | "line"
  | "bar"
  | "area"
  | "pie"
  | "scatter"
  | "doughnut"
  | "bubble"
  | "radar"
  | "linemarkers";

export const CHART_TYPES: readonly ChartType[] = [
  "column",
  "line",
  "bar",
  "area",
  "pie",
  "scatter",
  "doughnut",
  "bubble",
  "radar",
  "linemarkers",
];

export function isChartType(value: unknown): value is ChartType {
  return typeof value === "string" && (CHART_TYPES as readonly string[]).includes(value);
}

export interface ChartInfo {
  name: string;
  sheetName: string;
  chartType: ChartType | string;
  title?: string;
  titleVisible?: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  /** Office.js Chart.style (chart style id). */
  style?: number;
  /** Office.js Chart.legend.visible. */
  legendVisible?: boolean;
}

export interface ChartUpdateInput {
  sheetName: string;
  chartName: string;
  newName?: string;
  chartType?: ChartType;
  title?: string;
  showTitle?: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  /** Positive integer chart style id. */
  style?: number;
  /** Maps to Chart.legend.visible. */
  showLegend?: boolean;
}
