import type {
  ChartAxisDisplayUnit,
  ChartAxisGroup,
  ChartAxisInfo,
  ChartAxisKind,
  ChartAxisPosition,
  ChartAxisScaleType,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartAxisUpdateInput,
} from "./chartAxisTypes";

export const REQUIREMENT_SET = "ExcelApi";
export const EVIDENCE_17 =
  "ChartAxis displayUnit/scale/tickMark/tickLabelPosition require ExcelApi 1.7";
export const EVIDENCE_18 = "ChartAxis.position/setPositionAt require ExcelApi 1.8";
export const EVIDENCE_19 = "ChartAxis.linkNumberFormat requires ExcelApi 1.9";

export interface ExcelAxisTitle {
  text: string;
  visible: boolean;
  load(props: string): void;
}
export interface ExcelGridlines {
  visible: boolean;
  load(props: string): void;
}
export interface ExcelChartAxis {
  type?: string;
  axisGroup?: string;
  minimum: number | string | null;
  maximum: number | string | null;
  majorUnit: number | string | null;
  minorUnit: number | string | null;
  numberFormat: string | null;
  reversePlotOrder: boolean | null;
  displayUnit: string | null;
  customDisplayUnit: number | null;
  scaleType: string | null;
  logBase: number | null;
  showDisplayUnitLabel: boolean | null;
  majorTickMark: string | null;
  minorTickMark: string | null;
  tickLabelPosition: string | null;
  position: string | null;
  positionAt: number | null;
  linkNumberFormat: boolean | null;
  title: ExcelAxisTitle;
  majorGridlines: ExcelGridlines;
  minorGridlines: ExcelGridlines;
  setCustomDisplayUnit?(value: number): void;
  setPositionAt?(value: number): void;
  load(props: string): void;
}
export interface ExcelChartAxes {
  getItem(type: string, group?: string): ExcelChartAxis;
}

export const KIND_OFFICE: Record<ChartAxisKind, string> = {
  category: "Category",
  value: "Value",
};
export const GROUP_OFFICE: Record<ChartAxisGroup, string> = {
  primary: "Primary",
  secondary: "Secondary",
};
export const DISPLAY_UNIT_TO_HOST: Record<ChartAxisDisplayUnit, string> = {
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
export const SCALE_TO_HOST: Record<ChartAxisScaleType, string> = {
  linear: "Linear",
  logarithmic: "Logarithmic",
};
export const TICK_MARK_TO_HOST: Record<ChartAxisTickMark, string> = {
  none: "None",
  cross: "Cross",
  inside: "Inside",
  outside: "Outside",
};
export const TICK_LABEL_POS_TO_HOST: Record<ChartAxisTickLabelPosition, string> = {
  nextToAxis: "NextToAxis",
  high: "High",
  low: "Low",
  none: "None",
};
export const POSITION_TO_HOST: Record<ChartAxisPosition, string> = {
  automatic: "Automatic",
  maximum: "Maximum",
  minimum: "Minimum",
  custom: "Custom",
};

export function isSetSupported(version: string): boolean {
  const office = (globalThis as unknown as {
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office;
  const fn = office?.context?.requirements?.isSetSupported;
  if (typeof fn !== "function") return false;
  try {
    return fn.call(office!.context!.requirements, REQUIREMENT_SET, version);
  } catch {
    return false;
  }
}

export function mapKindLabel(raw: unknown): ChartAxisKind | string {
  if (typeof raw !== "string") throw new Error("ChartAxis.type is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "category" || key === "categories") return "category";
  if (key === "value" || key === "values") return "value";
  return raw;
}
export function mapGroupLabel(raw: unknown): ChartAxisGroup | string {
  if (typeof raw !== "string") throw new Error("ChartAxis.axisGroup is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "primary") return "primary";
  if (key === "secondary") return "secondary";
  return raw;
}
export function mapDisplayUnit(raw: unknown): ChartAxisDisplayUnit | string | null {
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
export function mapScaleType(raw: unknown): ChartAxisScaleType | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error("ChartAxis.scaleType is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "linear") return "linear";
  if (key === "logarithmic") return "logarithmic";
  return raw;
}
export function mapTickMark(raw: unknown, field: string): ChartAxisTickMark | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error(`${field} is not a loaded string`);
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartAxisTickMark> = {
    none: "none",
    cross: "cross",
    inside: "inside",
    outside: "outside",
  };
  return table[key] ?? raw;
}
export function mapTickLabelPosition(raw: unknown): ChartAxisTickLabelPosition | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error("ChartAxis.tickLabelPosition is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartAxisTickLabelPosition> = {
    nexttoaxis: "nextToAxis",
    high: "high",
    low: "low",
    none: "none",
  };
  return table[key] ?? raw;
}
export function mapPosition(raw: unknown): ChartAxisPosition | string | null {
  if (raw === null) return null;
  if (typeof raw !== "string") throw new Error("ChartAxis.position is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartAxisPosition> = {
    automatic: "automatic",
    maximum: "maximum",
    minimum: "minimum",
    custom: "custom",
  };
  return table[key] ?? raw;
}

export function readScalar(value: unknown, field: string): number | string | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error(`${field} is not a finite number`);
    return value;
  }
  if (typeof value === "string") return value;
  throw new Error(`${field} has invalid loaded type`);
}
export function readNumber(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "number" && Number.isFinite(value)) return value;
  throw new Error(`${field} has invalid loaded type`);
}
export function readBoolean(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "boolean") return value;
  throw new Error(`${field} has invalid loaded type`);
}
export function readString(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "string") return value;
  throw new Error(`${field} has invalid loaded type`);
}
export function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

export function needsExcelApi17(input: ChartAxisUpdateInput): boolean {
  return (
    input.displayUnit !== undefined ||
    input.customDisplayUnit !== undefined ||
    input.scaleType !== undefined ||
    input.logBase !== undefined ||
    input.showDisplayUnitLabel !== undefined ||
    input.majorTickMark !== undefined ||
    input.minorTickMark !== undefined ||
    input.tickLabelPosition !== undefined
  );
}
export function needsExcelApi18(input: ChartAxisUpdateInput): boolean {
  return input.position !== undefined || input.positionAt !== undefined;
}
export function needsExcelApi19(input: ChartAxisUpdateInput): boolean {
  return input.linkNumberFormat !== undefined;
}

export function applyAxisWrites(axis: ExcelChartAxis, input: ChartAxisUpdateInput): void {
  if (input.title !== undefined) {
    axis.title.text = input.title;
    axis.title.visible = input.title !== "";
  }
  if (input.minimum !== undefined) axis.minimum = input.minimum;
  if (input.maximum !== undefined) axis.maximum = input.maximum;
  if (input.majorUnit !== undefined) axis.majorUnit = input.majorUnit;
  if (input.minorUnit !== undefined) axis.minorUnit = input.minorUnit;
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
  if (input.majorTickMark !== undefined) {
    axis.majorTickMark = TICK_MARK_TO_HOST[input.majorTickMark];
  }
  if (input.minorTickMark !== undefined) {
    axis.minorTickMark = TICK_MARK_TO_HOST[input.minorTickMark];
  }
  if (input.tickLabelPosition !== undefined) {
    axis.tickLabelPosition = TICK_LABEL_POS_TO_HOST[input.tickLabelPosition];
  }
  if (input.position !== undefined) {
    if (input.position === "custom") {
      if (input.positionAt === undefined) {
        throw new Error("positionAt is required when position is custom");
      }
      if (typeof axis.setPositionAt !== "function") {
        throw new Error("setPositionAt missing (ExcelApi 1.8 required)");
      }
      axis.position = POSITION_TO_HOST.custom;
      axis.setPositionAt(input.positionAt);
    } else {
      axis.position = POSITION_TO_HOST[input.position];
    }
  } else if (input.positionAt !== undefined) {
    if (typeof axis.setPositionAt !== "function") {
      throw new Error("setPositionAt missing (ExcelApi 1.8 required)");
    }
    axis.setPositionAt(input.positionAt);
  }
  if (input.linkNumberFormat !== undefined) {
    axis.linkNumberFormat = input.linkNumberFormat;
  }
}

export function toAxisInfo(
  axis: ExcelChartAxis,
  sheetName: string,
  chartName: string,
  flags: { v17: boolean; v18: boolean; v19: boolean },
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
    minorUnit: readScalar(axis.minorUnit, "ChartAxis.minorUnit"),
    numberFormat: readString(axis.numberFormat, "ChartAxis.numberFormat"),
    reverse: readBoolean(axis.reversePlotOrder, "ChartAxis.reversePlotOrder"),
    displayUnit: flags.v17 ? mapDisplayUnit(axis.displayUnit) : null,
    customDisplayUnit: flags.v17
      ? readNumber(axis.customDisplayUnit, "ChartAxis.customDisplayUnit")
      : null,
    scaleType: flags.v17 ? mapScaleType(axis.scaleType) : null,
    logBase: flags.v17 ? readNumber(axis.logBase, "ChartAxis.logBase") : null,
    showDisplayUnitLabel: flags.v17
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
    majorTickMark: flags.v17
      ? mapTickMark(axis.majorTickMark, "ChartAxis.majorTickMark")
      : null,
    minorTickMark: flags.v17
      ? mapTickMark(axis.minorTickMark, "ChartAxis.minorTickMark")
      : null,
    tickLabelPosition: flags.v17 ? mapTickLabelPosition(axis.tickLabelPosition) : null,
    position: flags.v18 ? mapPosition(axis.position) : null,
    positionAt: flags.v18 ? readNumber(axis.positionAt, "ChartAxis.positionAt") : null,
    linkNumberFormat: flags.v19
      ? readBoolean(axis.linkNumberFormat, "ChartAxis.linkNumberFormat")
      : null,
  };
}
