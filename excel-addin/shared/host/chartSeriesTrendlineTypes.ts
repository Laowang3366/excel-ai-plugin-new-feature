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
  /** 1-based index within the series trendline collection (insertion order). */
  trendlineIndex: number;
  type: ChartTrendlineType | string;
  name: string | null;
  intercept: number | string | null;
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
  intercept?: number;
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
  intercept?: number;
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
