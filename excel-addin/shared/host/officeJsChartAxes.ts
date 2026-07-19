import { withExcel } from "./officeJsRuntime";
import type {
  ChartAxisGroup,
  ChartAxisInfo,
  ChartAxisKind,
  ChartAxisUpdateInput,
} from "./chartAxisTypes";
import type { HostResult } from "./types";

interface ExcelAxisTitle {
  text: string;
  visible: boolean;
  load(props: string): void;
}

interface ExcelChartAxis {
  type?: string;
  axisGroup?: string;
  minimum: number | string | null;
  maximum: number | string | null;
  majorUnit: number | string | null;
  numberFormat: string | null;
  reversePlotOrder: boolean | null;
  title: ExcelAxisTitle;
  load(props: string): void;
}

interface ExcelChartAxes {
  getItem(type: string, group?: string): ExcelChartAxis;
}

const KIND_OFFICE: Record<ChartAxisKind, string> = {
  category: "Category",
  value: "Value",
};

const GROUP_OFFICE: Record<ChartAxisGroup, string> = {
  primary: "Primary",
  secondary: "Secondary",
};

function mapKindLabel(raw: unknown): ChartAxisKind | string {
  if (typeof raw !== "string") {
    throw new Error("ChartAxis.type is not a loaded string");
  }
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "category" || key === "categories") return "category";
  if (key === "value" || key === "values") return "value";
  return raw;
}

function mapGroupLabel(raw: unknown): ChartAxisGroup | string {
  if (typeof raw !== "string") {
    throw new Error("ChartAxis.axisGroup is not a loaded string");
  }
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "primary") return "primary";
  if (key === "secondary") return "secondary";
  return raw;
}

/** number | string | null only; undefined/NaN/Infinity/object throw. */
function readScalar(value: unknown, field: string): number | string | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error(`${field} is not a finite number`);
    return value;
  }
  if (typeof value === "string") return value;
  throw new Error(`${field} has invalid loaded type`);
}

function readBoolean(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "boolean") return value;
  throw new Error(`${field} has invalid loaded type`);
}

function readString(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "string") return value;
  throw new Error(`${field} has invalid loaded type`);
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function toAxisInfo(axis: ExcelChartAxis, sheetName: string, chartName: string): ChartAxisInfo {
  return {
    sheetName,
    chartName,
    kind: mapKindLabel(axis.type),
    group: mapGroupLabel(axis.axisGroup),
    title: readString(axis.title.text, "ChartAxis.title.text"),
    titleVisible: readBoolean(axis.title.visible, "ChartAxis.title.visible"),
    minimum: readScalar(axis.minimum, "ChartAxis.minimum"),
    maximum: readScalar(axis.maximum, "ChartAxis.maximum"),
    majorUnit: readScalar(axis.majorUnit, "ChartAxis.majorUnit"),
    numberFormat: readString(axis.numberFormat, "ChartAxis.numberFormat"),
    reverse: readBoolean(axis.reversePlotOrder, "ChartAxis.reversePlotOrder"),
  };
}

/** Update chart axis fields; write → sync → load+sync real snapshot. */
export async function officeJsUpdateChartAxis(
  input: ChartAxisUpdateInput,
): Promise<HostResult<ChartAxisInfo>> {
  return withExcel("chart.axes.update", async (context) => {
    const group: ChartAxisGroup = input.group ?? "primary";
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      axes: ExcelChartAxes;
      load(props: string): void;
    };
    const axis = chart.axes.getItem(KIND_OFFICE[input.kind], GROUP_OFFICE[group]);

    if (input.title !== undefined) {
      axis.title.text = input.title;
      axis.title.visible = input.title !== "";
    }
    if (input.minimum !== undefined) axis.minimum = input.minimum;
    if (input.maximum !== undefined) axis.maximum = input.maximum;
    if (input.majorUnit !== undefined) axis.majorUnit = input.majorUnit;
    if (input.numberFormat !== undefined) axis.numberFormat = input.numberFormat;
    if (input.reverse !== undefined) axis.reversePlotOrder = input.reverse;

    await context.sync();
    chart.load("name");
    axis.load("type,axisGroup,minimum,maximum,majorUnit,numberFormat,reversePlotOrder");
    axis.title.load("text,visible");
    await context.sync();
    return toAxisInfo(axis, input.sheetName, requireLoadedString(chart.name, "Chart.name"));
  });
}
