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
}

/**
 * Host snapshot after write→sync→load→sync.
 * enabled always present (hasDataLabels).
 * show fields and numberFormat present only on ExcelApi 1.8 path (dataLabels fields touched).
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
}
