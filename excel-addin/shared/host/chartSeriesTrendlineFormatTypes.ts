/** Public ChartLineStyle labels (host PascalCase). ExcelApi 1.7. */
export const CHART_LINE_STYLES = [
  "none",
  "continuous",
  "dash",
  "dashDot",
  "dashDotDot",
  "dot",
  "grey25",
  "grey50",
  "grey75",
  "automatic",
  "roundDot",
] as const;

export type ChartLineStyle = (typeof CHART_LINE_STYLES)[number];

export function isChartLineStyle(value: unknown): value is ChartLineStyle {
  return typeof value === "string" && (CHART_LINE_STYLES as readonly string[]).includes(value);
}

/** Write input for chart.series.trendlines.format.update (ExcelApi 1.7). */
export interface ChartTrendlineFormatUpdateInput {
  sheetName: string;
  chartName: string;
  /** Public 1-based series index. */
  seriesIndex: number;
  /** Public 1-based trendline index (host getItem is 0-based). */
  trendlineIndex: number;
  /** #RRGGBB only after host normalization. */
  color?: string;
  lineStyle?: ChartLineStyle;
  /** Weight in points; finite number. No official min/max in Office.js docs; host may reject invalid values. */
  weight?: number;
}

/** Host readback snapshot of trendline line format. */
export interface ChartTrendlineFormatInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  trendlineIndex: number;
  color: string;
  lineStyle: ChartLineStyle | string;
  weight: number;
}
