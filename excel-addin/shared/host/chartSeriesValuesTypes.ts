/** Input for chart.series.values.update (ExcelApi 1.15 source readback). */
export interface ChartSeriesValuesUpdateInput {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  valuesRange?: string;
  xValuesRange?: string;
}

/**
 * Host-verified series data-source snapshot.
 * valuesSource / xValuesSource come from getDimensionDataSourceString, not input echo.
 */
export interface ChartSeriesValuesInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  valuesSource?: string;
  xValuesSource?: string;
  dataBound: true;
}
