/** Input for chart.series.bubbleSizes.update (ExcelApi 1.15 BubbleSizes readback). */
export interface ChartSeriesBubbleSizesUpdateInput {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  bubbleSizesRange: string;
}

/**
 * Host-verified bubble-sizes source snapshot.
 * bubbleSizesSource comes from getDimensionDataSourceString("BubbleSizes"), not input echo.
 * sheetName/chartName come from loaded worksheet.name / chart.name.
 */
export interface ChartSeriesBubbleSizesInfo {
  sheetName: string;
  chartName: string;
  seriesIndex: number;
  bubbleSizesSource: string;
  dataBound: true;
}
