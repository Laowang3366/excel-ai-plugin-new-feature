import type { HostResult } from "./types";
import { unsupported } from "./types";

const CHART_SERIES_EVIDENCE = "No in-repo WPS JSA ChartSeries contract";

export async function wpsListChartSeries(_sheetName: string, _chartName: string) {
  return unsupported(
    "chart.series.list",
    "wps-jsa",
    "chart.series.list is not verified for WPS JSA",
    CHART_SERIES_EVIDENCE,
  ) as HostResult<never>;
}

export async function wpsUpdateChartSeries(_input: unknown) {
  return unsupported(
    "chart.series.update",
    "wps-jsa",
    "chart.series.update is not verified for WPS JSA",
    CHART_SERIES_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SET_DATA_EVIDENCE = "No in-repo WPS JSA Chart.setData contract";

export async function wpsUpdateChartSource(_input: unknown) {
  return unsupported(
    "chart.source.update",
    "wps-jsa",
    "chart.source.update is not verified for WPS JSA",
    CHART_SET_DATA_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_AXIS_EVIDENCE = "No in-repo WPS JSA ChartAxis contract";

export async function wpsUpdateChartAxis(_input: unknown) {
  return unsupported(
    "chart.axes.update",
    "wps-jsa",
    "chart.axes.update is not verified for WPS JSA",
    CHART_AXIS_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_DATA_LABELS_EVIDENCE =
  "No in-repo WPS JSA ChartSeries.hasDataLabels/dataLabels.enabled contract";

export async function wpsUpdateChartDataLabels(_input: unknown) {
  return unsupported(
    "chart.series.dataLabels.update",
    "wps-jsa",
    "chart.series.dataLabels.update is not verified for WPS JSA",
    CHART_DATA_LABELS_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_AXIS_GROUP_EVIDENCE = "No in-repo WPS JSA ChartSeries.axisGroup contract";

export async function wpsUpdateChartSeriesAxisGroup(_input: unknown) {
  return unsupported(
    "chart.series.axisGroup.update",
    "wps-jsa",
    "chart.series.axisGroup.update is not verified for WPS JSA",
    CHART_SERIES_AXIS_GROUP_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_DELETE_EVIDENCE = "No in-repo WPS JSA ChartSeries.delete contract";

export async function wpsDeleteChartSeries(
  _sheetName: string,
  _chartName: string,
  _seriesIndex: number,
) {
  return unsupported(
    "chart.series.delete",
    "wps-jsa",
    "chart.series.delete is not verified for WPS JSA",
    CHART_SERIES_DELETE_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_ADD_EVIDENCE = "No in-repo WPS JSA ChartSeriesCollection.add contract";

export async function wpsAddChartSeries(_input: unknown) {
  return unsupported(
    "chart.series.add",
    "wps-jsa",
    "chart.series.add is not verified for WPS JSA",
    CHART_SERIES_ADD_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_VALUES_EVIDENCE =
  "No in-repo WPS JSA ChartSeries values/xValues contract";

export async function wpsUpdateChartSeriesValues(_input: unknown) {
  return unsupported(
    "chart.series.values.update",
    "wps-jsa",
    "chart.series.values.update is not verified for WPS JSA",
    CHART_SERIES_VALUES_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_BUBBLE_SIZES_EVIDENCE =
  "No in-repo WPS JSA ChartSeries.setBubbleSizes/BubbleSizes contract";

export async function wpsUpdateChartSeriesBubbleSizes(_input: unknown) {
  return unsupported(
    "chart.series.bubbleSizes.update",
    "wps-jsa",
    "chart.series.bubbleSizes.update is not verified for WPS JSA",
    CHART_SERIES_BUBBLE_SIZES_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_IMAGE_EVIDENCE = "No in-repo WPS JSA Chart.getImage/export contract";

export async function wpsGetChartImage(_input: unknown) {
  return unsupported(
    "chart.image.get",
    "wps-jsa",
    "chart.image.get is not verified for WPS JSA",
    CHART_IMAGE_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_TRENDLINES_EVIDENCE =
  "No in-repo WPS JSA ChartSeries.trendlines / ChartTrendline contract";

export async function wpsListChartSeriesTrendlines(
  _sheetName: string,
  _chartName: string,
  _seriesIndex: number,
) {
  return unsupported(
    "chart.series.trendlines.list",
    "wps-jsa",
    "chart.series.trendlines.list is not verified for WPS JSA",
    CHART_SERIES_TRENDLINES_EVIDENCE,
  ) as HostResult<never>;
}

export async function wpsAddChartSeriesTrendline(_input: unknown) {
  return unsupported(
    "chart.series.trendlines.add",
    "wps-jsa",
    "chart.series.trendlines.add is not verified for WPS JSA",
    CHART_SERIES_TRENDLINES_EVIDENCE,
  ) as HostResult<never>;
}

export async function wpsUpdateChartSeriesTrendline(_input: unknown) {
  return unsupported(
    "chart.series.trendlines.update",
    "wps-jsa",
    "chart.series.trendlines.update is not verified for WPS JSA",
    CHART_SERIES_TRENDLINES_EVIDENCE,
  ) as HostResult<never>;
}

export async function wpsDeleteChartSeriesTrendline(
  _sheetName: string,
  _chartName: string,
  _seriesIndex: number,
  _trendlineIndex: number,
) {
  return unsupported(
    "chart.series.trendlines.delete",
    "wps-jsa",
    "chart.series.trendlines.delete is not verified for WPS JSA",
    CHART_SERIES_TRENDLINES_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_MARKERS_EVIDENCE =
  "No in-repo WPS JSA ChartSeries.markerStyle/markerSize/marker color contract";

export async function wpsUpdateChartSeriesMarkers(_input: unknown) {
  return unsupported(
    "chart.series.markers.update",
    "wps-jsa",
    "chart.series.markers.update is not verified for WPS JSA",
    CHART_SERIES_MARKERS_EVIDENCE,
  ) as HostResult<never>;
}

const CHART_SERIES_TRENDLINE_FORMAT_EVIDENCE =
  "No in-repo WPS JSA ChartTrendline.format.line contract";

export async function wpsUpdateChartSeriesTrendlineFormat(_input: unknown) {
  return unsupported(
    "chart.series.trendlines.format.update",
    "wps-jsa",
    "chart.series.trendlines.format.update is not verified for WPS JSA",
    CHART_SERIES_TRENDLINE_FORMAT_EVIDENCE,
  ) as HostResult<never>;
}
