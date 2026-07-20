export type ChartAxisKind = "category" | "value";
export type ChartAxisGroup = "primary" | "secondary";
export type ChartAxisScaleType = "linear" | "logarithmic";

/** Public labels; host tokens are PascalCase (None, Hundreds, …). Custom uses setCustomDisplayUnit. */
export type ChartAxisDisplayUnit =
  | "none"
  | "hundreds"
  | "thousands"
  | "tenThousands"
  | "hundredThousands"
  | "millions"
  | "tenMillions"
  | "hundredMillions"
  | "billions"
  | "trillions"
  | "custom";

export const CHART_AXIS_KINDS: readonly ChartAxisKind[] = ["category", "value"];
export const CHART_AXIS_GROUPS: readonly ChartAxisGroup[] = ["primary", "secondary"];
export const CHART_AXIS_SCALE_TYPES: readonly ChartAxisScaleType[] = ["linear", "logarithmic"];
export const CHART_AXIS_DISPLAY_UNITS: readonly ChartAxisDisplayUnit[] = [
  "none",
  "hundreds",
  "thousands",
  "tenThousands",
  "hundredThousands",
  "millions",
  "tenMillions",
  "hundredMillions",
  "billions",
  "trillions",
  "custom",
];

export function isChartAxisKind(value: unknown): value is ChartAxisKind {
  return typeof value === "string" && (CHART_AXIS_KINDS as readonly string[]).includes(value);
}

export function isChartAxisGroup(value: unknown): value is ChartAxisGroup {
  return typeof value === "string" && (CHART_AXIS_GROUPS as readonly string[]).includes(value);
}

export function isChartAxisScaleType(value: unknown): value is ChartAxisScaleType {
  return typeof value === "string" && (CHART_AXIS_SCALE_TYPES as readonly string[]).includes(value);
}

export function isChartAxisDisplayUnit(value: unknown): value is ChartAxisDisplayUnit {
  return typeof value === "string" && (CHART_AXIS_DISPLAY_UNITS as readonly string[]).includes(value);
}

export interface ChartAxisUpdateInput {
  sheetName: string;
  chartName: string;
  kind: ChartAxisKind;
  /** Defaults to primary. */
  group?: ChartAxisGroup;
  /** Empty string clears/hides title; non-empty must be trim-valid. */
  title?: string;
  minimum?: number;
  maximum?: number;
  /** Finite number >= 0. */
  majorUnit?: number;
  numberFormat?: string;
  reverse?: boolean;
  /** ExcelApi 1.7 ChartAxis.displayUnit (Custom requires customDisplayUnit). */
  displayUnit?: ChartAxisDisplayUnit;
  /** ExcelApi 1.7 setCustomDisplayUnit; required when displayUnit is custom. */
  customDisplayUnit?: number;
  /** ExcelApi 1.7 ChartAxis.scaleType. */
  scaleType?: ChartAxisScaleType;
  /** ExcelApi 1.7 ChartAxis.logBase (meaningful with logarithmic). */
  logBase?: number;
  /** ExcelApi 1.7 ChartAxis.showDisplayUnitLabel. */
  showDisplayUnitLabel?: boolean;
  /** ExcelApi 1.1 ChartAxis.majorGridlines.visible. */
  majorGridlinesVisible?: boolean;
  /** ExcelApi 1.1 ChartAxis.minorGridlines.visible. */
  minorGridlinesVisible?: boolean;
}

/** Real axis snapshot after write→sync→load→sync (not input echo). */
export interface ChartAxisInfo {
  sheetName: string;
  chartName: string;
  /** Mapped category|value or raw host string. */
  kind: ChartAxisKind | string;
  /** Mapped primary|secondary or raw host string. */
  group: ChartAxisGroup | string;
  title: string | null;
  titleVisible: boolean | null;
  /** Host scale may be number, string auto, or null. */
  minimum: number | string | null;
  maximum: number | string | null;
  majorUnit: number | string | null;
  numberFormat: string | null;
  reverse: boolean | null;
  displayUnit: ChartAxisDisplayUnit | string | null;
  customDisplayUnit: number | null;
  scaleType: ChartAxisScaleType | string | null;
  logBase: number | null;
  showDisplayUnitLabel: boolean | null;
  majorGridlinesVisible: boolean | null;
  minorGridlinesVisible: boolean | null;
}
