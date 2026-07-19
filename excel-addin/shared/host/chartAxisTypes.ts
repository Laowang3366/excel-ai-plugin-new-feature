export type ChartAxisKind = "category" | "value";
export type ChartAxisGroup = "primary" | "secondary";

export const CHART_AXIS_KINDS: readonly ChartAxisKind[] = ["category", "value"];
export const CHART_AXIS_GROUPS: readonly ChartAxisGroup[] = ["primary", "secondary"];

export function isChartAxisKind(value: unknown): value is ChartAxisKind {
  return typeof value === "string" && (CHART_AXIS_KINDS as readonly string[]).includes(value);
}

export function isChartAxisGroup(value: unknown): value is ChartAxisGroup {
  return typeof value === "string" && (CHART_AXIS_GROUPS as readonly string[]).includes(value);
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
}
