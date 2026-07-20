export type ChartDataLabelPosition =
  | "none"
  | "center"
  | "insideEnd"
  | "insideBase"
  | "outsideEnd"
  | "left"
  | "right"
  | "top"
  | "bottom"
  | "bestFit"
  | "callout";

export const CHART_DATA_LABEL_POSITIONS: readonly ChartDataLabelPosition[] = [
  "none",
  "center",
  "insideEnd",
  "insideBase",
  "outsideEnd",
  "left",
  "right",
  "top",
  "bottom",
  "bestFit",
  "callout",
] as const;

export function isChartDataLabelPosition(value: unknown): value is ChartDataLabelPosition {
  return (
    typeof value === "string" &&
    (CHART_DATA_LABEL_POSITIONS as readonly string[]).includes(value)
  );
}

export interface ChartDataLabelsUpdateInput {
  sheetName: string;
  chartName: string;
  /** Public 1-based series index. */
  seriesIndex: number;
  /** Maps to ChartSeries.hasDataLabels (ExcelApi 1.7). */
  enabled?: boolean;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  /** Non-empty after trim. */
  numberFormat?: string;
  /** ExcelApi 1.8 path via ChartSeries.dataLabels (member ExcelApi 1.1). */
  showPercentage?: boolean;
  showBubbleSize?: boolean;
  showLegendKey?: boolean;
  /**
   * ChartDataLabels.separator — preserved as-is (including "" and edge spaces).
   * Official contract has no non-empty/length limit.
   */
  separator?: string;
  /** ChartDataLabelPosition without Invalid. */
  position?: ChartDataLabelPosition;
}

/**
 * Host snapshot after write→sync→load→sync.
 * enabled always present (hasDataLabels).
 * show/position/separator/numberFormat present only on ExcelApi 1.8 path (dataLabels fields touched).
 * enabled-only on 1.7 omits those fields rather than fabricating them.
 */
export interface ChartDataLabelsInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  /** Host ChartSeries.hasDataLabels readback. */
  enabled: boolean;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  numberFormat?: string;
  showPercentage?: boolean;
  showBubbleSize?: boolean;
  showLegendKey?: boolean;
  separator?: string;
  position?: ChartDataLabelPosition | string;
}
