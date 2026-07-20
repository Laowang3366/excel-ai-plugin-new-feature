/** Public ChartSeries marker style (Office.js ChartMarkerStyle without Invalid write). */
export const CHART_MARKER_STYLES = [
  "automatic",
  "none",
  "square",
  "diamond",
  "triangle",
  "x",
  "star",
  "dot",
  "dash",
  "circle",
  "plus",
  "picture",
] as const;

export type ChartMarkerStyle = (typeof CHART_MARKER_STYLES)[number];

export function isChartMarkerStyle(value: unknown): value is ChartMarkerStyle {
  return typeof value === "string" && (CHART_MARKER_STYLES as readonly string[]).includes(value);
}

/** Write input for chart.series.markers.update (ExcelApi 1.7). */
export interface ChartSeriesMarkersUpdateInput {
  sheetName: string;
  chartName: string;
  /** Public 1-based series index. */
  seriesIndex: number;
  markerStyle?: ChartMarkerStyle;
  /** Host range 2..72. */
  markerSize?: number;
  /** #RRGGBB only. */
  markerBackgroundColor?: string;
  /** #RRGGBB only. */
  markerForegroundColor?: string;
}

/** Host readback snapshot after markers write. */
export interface ChartSeriesMarkersInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  markerStyle: ChartMarkerStyle | string;
  markerSize: number;
  markerBackgroundColor: string;
  markerForegroundColor: string;
}
