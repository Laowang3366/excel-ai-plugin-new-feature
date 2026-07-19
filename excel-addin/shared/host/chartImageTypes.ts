/** Input for chart.image.get (ExcelApi 1.2 Chart.getImage). */
export interface ChartImageGetInput {
  sheetName: string;
  chartName: string;
  /** Optional desired image width (1–4096). */
  width?: number;
  /** Optional desired image height (1–4096). */
  height?: number;
}

/** Host-generated chart image snapshot (Base64 payload only; no path/MIME claim). */
export interface ChartImageInfo {
  sheetName: string;
  chartName: string;
  imageBase64: string;
}
