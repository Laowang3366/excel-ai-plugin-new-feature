/** Public ChartTrendline.type labels (host PascalCase). ExcelApi 1.7. */
export type ChartTrendlineType =
  | "linear"
  | "exponential"
  | "logarithmic"
  | "movingAverage"
  | "polynomial"
  | "power";

export const CHART_TRENDLINE_TYPES: readonly ChartTrendlineType[] = [
  "linear",
  "exponential",
  "logarithmic",
  "movingAverage",
  "polynomial",
  "power",
];

export function isChartTrendlineType(value: unknown): value is ChartTrendlineType {
  return typeof value === "string" && (CHART_TRENDLINE_TYPES as readonly string[]).includes(value);
}

export interface ChartTrendlineInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  /**
   * 1-based public index (insertion order).
   * Host `ChartTrendlineCollection.getItem` is 0-based.
   */
  trendlineIndex: number;
  type: ChartTrendlineType | string;
  /** Host readback is always a string per Office.js. */
  name: string | null;
  /**
   * Host readback is always a number per Office.js.
   * Writes may use "" for automatic intercept.
   */
  intercept: number | null;
  polynomialOrder: number | null;
  movingAveragePeriod: number | null;
  /** ExcelApi 1.8; null when requirement set not available. */
  forwardPeriod: number | null;
  backwardPeriod: number | null;
  showEquation: boolean | null;
  showRSquared: boolean | null;
}

export interface ChartTrendlineListResult {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  trendlines: ChartTrendlineInfo[];
}

export interface ChartTrendlineAddInput {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  type: ChartTrendlineType;
  name?: string;
  /** Finite number, or empty string for automatic (Office.js write contract). */
  intercept?: number | "";
  polynomialOrder?: number;
  movingAveragePeriod?: number;
  forwardPeriod?: number;
  backwardPeriod?: number;
  showEquation?: boolean;
  showRSquared?: boolean;
}

export interface ChartTrendlineUpdateInput {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  trendlineIndex: number;
  type?: ChartTrendlineType;
  name?: string;
  /** Finite number, or empty string for automatic (Office.js write contract). */
  intercept?: number | "";
  polynomialOrder?: number;
  movingAveragePeriod?: number;
  forwardPeriod?: number;
  backwardPeriod?: number;
  showEquation?: boolean;
  showRSquared?: boolean;
}

export interface ChartTrendlineDeleteResult {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  deletedTrendlineIndex: number;
  remainingTrendlines: ChartTrendlineInfo[];
}
