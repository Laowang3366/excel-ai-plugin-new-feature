import type { ChartType } from "./chartTypes";

export function mapChartType(type: ChartType | undefined): string {
  const chartTypes = typeof window !== "undefined" ? window.Excel?.ChartType : undefined;
  switch (type) {
    case "line":
      return chartTypes?.line ?? "Line";
    case "bar":
      return chartTypes?.barClustered ?? "BarClustered";
    case "area":
      return chartTypes?.area ?? "Area";
    case "pie":
      return chartTypes?.pie ?? "Pie";
    case "scatter":
      return chartTypes?.xyscatter ?? "XYScatter";
    case "doughnut":
      return chartTypes?.doughnut ?? "Doughnut";
    case "bubble":
      return chartTypes?.bubble ?? "Bubble";
    case "radar":
      return chartTypes?.radar ?? "Radar";
    case "linemarkers":
      return chartTypes?.lineMarkers ?? "LineMarkers";
    case "column":
    default:
      return chartTypes?.columnClustered ?? "ColumnClustered";
  }
}

/** Official exact host names for deep types only (no fuzzy/alias match). */
const EXACT_DEEP_CHART_TYPES: Readonly<Record<string, ChartType>> = {
  Doughnut: "doughnut",
  Bubble: "bubble",
  Radar: "radar",
  LineMarkers: "linemarkers",
};

/** Basic six + stacked/adjacent variants (unchanged Phase13 behavior). */
const KNOWN_CHART_TYPE_LABELS: Readonly<Record<string, ChartType>> = {
  columnclustered: "column",
  columnstacked: "column",
  columnstacked100: "column",
  line: "line",
  linestacked: "line",
  linestacked100: "line",
  linemarkersstacked: "line",
  linemarkersstacked100: "line",
  barclustered: "bar",
  barstacked: "bar",
  barstacked100: "bar",
  area: "area",
  areastacked: "area",
  areastacked100: "area",
  pie: "pie",
  pieexploded: "pie",
  pieofpie: "pie",
  barofpie: "pie",
  xyscatter: "scatter",
  xyscatterlines: "scatter",
  xyscatterlinesnomarkers: "scatter",
  xyscattersmooth: "scatter",
  xyscattersmoothnomarkers: "scatter",
  scatter: "scatter",
};

export function toChartTypeLabel(value: string): ChartType | string {
  const exact = EXACT_DEEP_CHART_TYPES[value];
  if (exact) return exact;
  const key = value.toLowerCase().replace(/[^a-z0-9]/g, "");
  return KNOWN_CHART_TYPE_LABELS[key] ?? value;
}
