import { withExcel } from "./officeJsRuntime";
import type {
  ChartAxisDisplayUnit,
  ChartAxisGroup,
  ChartAxisInfo,
  ChartAxisKind,
  ChartAxisScaleType,
  ChartAxisUpdateInput,
} from "./chartAxisTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQUIREMENT_VERSION = "1.7";
const REQUIREMENT_EVIDENCE =
  "ChartAxis.displayUnit/scaleType/showDisplayUnitLabel/logBase/setCustomDisplayUnit require ExcelApi 1.7";

interface ExcelAxisTitle {
  text: string;
  visible: boolean;
  load(props: string): void;
}

interface ExcelGridlines {
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
  displayUnit: string | null;
  customDisplayUnit: number | null;
  scaleType: string | null;
  logBase: number | null;
  showDisplayUnitLabel: boolean | null;
  title: ExcelAxisTitle;
  majorGridlines: ExcelGridlines;
  minorGridlines: ExcelGridlines;
  setCustomDisplayUnit?(value: number): void;
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

const DISPLAY_UNIT_TO_HOST: Record<ChartAxisDisplayUnit, string> = {
  none: "None",
  hundreds: "Hundreds",
  thousands: "Thousands",
  tenThousands: "TenThousands",
  hundredThousands: "HundredThousands",
  millions: "Millions",
  tenMillions: "TenMillions",
  hundredMillions: "HundredMillions",
  billions: "Billions",
  trillions: "Trillions",
  custom: "Custom",
};

const SCALE_TO_HOST: Record<ChartAxisScaleType, string> = {
  linear: "Linear",
  logarithmic: "Logarithmic",
};

function mapKindLabel(raw: unknown): ChartAxisKind | string {
  if (typeof raw !== "string") throw new Error("ChartAxis.type is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "category" || key === "categories") return "category";
  if (key === "value" || key === "values") return "value";
  return raw;
}

function mapGroupLabel(raw: unknown): ChartAxisGroup | string {
  if (typeof raw !== "string") throw new Error("ChartAxis.axisGroup is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "primary") return "primary";
  if (key === "secondary") return "secondary";
  return raw;
}

function mapDisplayUnit(raw: unknown): ChartAxisDisplayUnit | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error("ChartAxis.displayUnit is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartAxisDisplayUnit> = {
    none: "none",
    hundreds: "hundreds",
    thousands: "thousands",
    tenthousands: "tenThousands",
    hundredthousands: "hundredThousands",
    millions: "millions",
    tenmillions: "tenMillions",
    hundredmillions: "hundredMillions",
    billions: "billions",
    trillions: "trillions",
    custom: "custom",
  };
  return table[key] ?? raw;
}

function mapScaleType(raw: unknown): ChartAxisScaleType | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error("ChartAxis.scaleType is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "linear") return "linear";
  if (key === "logarithmic") return "logarithmic";
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

function readNumber(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "number" && Number.isFinite(value)) return value;
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

function needsExcelApi17(input: ChartAxisUpdateInput): boolean {
  return (
    input.displayUnit !== undefined ||
    input.customDisplayUnit !== undefined ||
    input.scaleType !== undefined ||
    input.logBase !== undefined ||
    input.showDisplayUnitLabel !== undefined
  );
}

/** Official precheck for ExcelApi 1.7 axis scale/display members. */
export function isExcelApi17Supported(): boolean {
  const office = (globalThis as unknown as {
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, REQUIREMENT_SET, REQUIREMENT_VERSION);
  } catch {
    return false;
  }
}

function toAxisInfo(
  axis: ExcelChartAxis,
  sheetName: string,
  chartName: string,
  includeAdvanced: boolean,
): ChartAxisInfo {
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
    displayUnit: includeAdvanced ? mapDisplayUnit(axis.displayUnit) : null,
    customDisplayUnit: includeAdvanced
      ? readNumber(axis.customDisplayUnit, "ChartAxis.customDisplayUnit")
      : null,
    scaleType: includeAdvanced ? mapScaleType(axis.scaleType) : null,
    logBase: includeAdvanced ? readNumber(axis.logBase, "ChartAxis.logBase") : null,
    showDisplayUnitLabel: includeAdvanced
      ? readBoolean(axis.showDisplayUnitLabel, "ChartAxis.showDisplayUnitLabel")
      : null,
    majorGridlinesVisible: readBoolean(
      axis.majorGridlines.visible,
      "ChartAxis.majorGridlines.visible",
    ),
    minorGridlinesVisible: readBoolean(
      axis.minorGridlines.visible,
      "ChartAxis.minorGridlines.visible",
    ),
  };
}

function applyWrites(axis: ExcelChartAxis, input: ChartAxisUpdateInput): void {
  if (input.title !== undefined) {
    axis.title.text = input.title;
    axis.title.visible = input.title !== "";
  }
  if (input.minimum !== undefined) axis.minimum = input.minimum;
  if (input.maximum !== undefined) axis.maximum = input.maximum;
  if (input.majorUnit !== undefined) axis.majorUnit = input.majorUnit;
  if (input.numberFormat !== undefined) axis.numberFormat = input.numberFormat;
  if (input.reverse !== undefined) axis.reversePlotOrder = input.reverse;
  if (input.displayUnit !== undefined) {
    if (input.displayUnit === "custom") {
      if (input.customDisplayUnit === undefined) {
        throw new Error("customDisplayUnit is required when displayUnit is custom");
      }
      if (typeof axis.setCustomDisplayUnit !== "function") {
        throw new Error("setCustomDisplayUnit missing (ExcelApi 1.7 required)");
      }
      axis.setCustomDisplayUnit(input.customDisplayUnit);
    } else {
      axis.displayUnit = DISPLAY_UNIT_TO_HOST[input.displayUnit];
    }
  } else if (input.customDisplayUnit !== undefined) {
    if (typeof axis.setCustomDisplayUnit !== "function") {
      throw new Error("setCustomDisplayUnit missing (ExcelApi 1.7 required)");
    }
    axis.setCustomDisplayUnit(input.customDisplayUnit);
  }
  if (input.scaleType !== undefined) axis.scaleType = SCALE_TO_HOST[input.scaleType];
  if (input.logBase !== undefined) axis.logBase = input.logBase;
  if (input.showDisplayUnitLabel !== undefined) {
    axis.showDisplayUnitLabel = input.showDisplayUnitLabel;
  }
  if (input.majorGridlinesVisible !== undefined) {
    axis.majorGridlines.visible = input.majorGridlinesVisible;
  }
  if (input.minorGridlinesVisible !== undefined) {
    axis.minorGridlines.visible = input.minorGridlinesVisible;
  }
}

/** Update chart axis fields; write → sync → load+sync real snapshot. */
export async function officeJsUpdateChartAxis(
  input: ChartAxisUpdateInput,
): Promise<HostResult<ChartAxisInfo>> {
  if (needsExcelApi17(input) && !isExcelApi17Supported()) {
    return unsupported(
      "chart.axes.update",
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  return withExcel("chart.axes.update", async (context) => {
    const group: ChartAxisGroup = input.group ?? "primary";
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      axes: ExcelChartAxes;
      load(props: string): void;
    };
    const axis = chart.axes.getItem(KIND_OFFICE[input.kind], GROUP_OFFICE[group]);
    applyWrites(axis, input);
    await context.sync();
    const includeAdvanced = isExcelApi17Supported();
    chart.load("name");
    const baseProps =
      "type,axisGroup,minimum,maximum,majorUnit,numberFormat,reversePlotOrder";
    axis.load(
      includeAdvanced
        ? `${baseProps},displayUnit,customDisplayUnit,scaleType,logBase,showDisplayUnitLabel`
        : baseProps,
    );
    axis.title.load("text,visible");
    axis.majorGridlines.load("visible");
    axis.minorGridlines.load("visible");
    await context.sync();
    return toAxisInfo(
      axis,
      input.sheetName,
      requireLoadedString(chart.name, "Chart.name"),
      includeAdvanced,
    );
  });
}
