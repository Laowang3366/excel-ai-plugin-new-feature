/**
 * Chart.setData data-source update (Office.js).
 * Evidence: Chart.setData(sourceData: Range, seriesBy?: ChartSeriesBy) takes a Range
 * proxy — not a sheet-local address string — so source may come from another worksheet
 * in the same workbook via worksheets.getItem(...).getRange(...).
 */
import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartSeriesInfo } from "./chartSeriesTypes";
import type {
  ChartSeriesBy,
  ChartSourceInfo,
  ChartSourceUpdateInput,
} from "./chartSourceTypes";
import type { HostResult } from "./types";

interface ExcelChartSeries {
  name: string;
  chartType: string;
  smooth: boolean;
}

interface ExcelChartWithSetData {
  name: string;
  setData?(range: object, seriesBy?: string): void;
  load(props: string): void;
  series: {
    items: ExcelChartSeries[];
    load(props: string): void;
  };
}

export type ParsedChartSourceRange = {
  /** Worksheet that owns the data range (may differ from chart sheet). */
  sourceSheetName: string;
  /** Uppercase bare A1 (no sheet qualifier). */
  bareA1: string;
  /** Canonical value returned to callers (bare when same sheet; qualified when cross-sheet). */
  displaySourceRange: string;
};

const SERIES_BY_OFFICE: Record<ChartSeriesBy, string> = {
  auto: "Auto",
  rows: "Rows",
  columns: "Columns",
};

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") {
    throw new Error(`${field} is not a loaded string`);
  }
  return value;
}

function requireLoadedBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a loaded boolean`);
  }
  return value;
}

function validateBareA1(bareRaw: string, fieldName: string): string {
  let bare = bareRaw.replace(/\$/g, "").trim();
  if (bare === "") throw new Error(`${fieldName} must be non-empty`);
  if (bare.includes(",")) {
    throw new Error(`${fieldName} multi-area is not supported`);
  }
  if (bare.includes("[") || bare.includes("]")) {
    throw new Error(`${fieldName} structured references are not supported`);
  }
  if (!/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(bare)) {
    throw new Error(`${fieldName} must be a single contiguous A1 address`);
  }
  bare = bare.toUpperCase();
  for (const part of bare.split(":")) {
    const rowMatch = /^[A-Za-z]+(\d+)$/.exec(part);
    if (!rowMatch || Number(rowMatch[1]) < 1) {
      throw new Error(`${fieldName} row must be >= 1`);
    }
  }
  return bare;
}

function formatSheetQualifiedA1(sheetName: string, bareA1: string, chartSheetName: string): string {
  if (sheetName.toLowerCase() === chartSheetName.toLowerCase()) return bareA1;
  const needsQuotes = /[\s'!]/.test(sheetName) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName);
  if (needsQuotes) return `'${sheetName.replace(/'/g, "''")}'!${bareA1}`;
  return `${sheetName}!${bareA1}`;
}

/**
 * Parse chart sourceRange: bare A1 (chart sheet) or Sheet!A1 / 'Sheet 2'!A1.
 * Rejects external workbooks, 3D refs, multi-area, and structured references.
 */
export function parseChartSourceRange(
  chartSheetName: string,
  sourceRange: string,
): ParsedChartSourceRange {
  const raw = sourceRange.trim();
  if (raw === "") throw new Error("sourceRange must be non-empty");

  // External workbook: [Book.xlsx]Sheet!A1 or '[Book.xlsx]Sheet'!A1
  if (
    raw.startsWith("[") ||
    raw.startsWith("'[") ||
    (/\[[^\]]+\]/.test(raw) && raw.includes("!"))
  ) {
    throw new Error("sourceRange external workbook references are not supported");
  }

  // Structured table refs (no sheet bang, or A1 part has brackets).
  if (!raw.includes("!") && (raw.includes("[") || raw.includes("]"))) {
    throw new Error("sourceRange structured references are not supported");
  }

  let sourceSheetName = chartSheetName;
  let a1Part = raw;

  if (raw.includes("!")) {
    if (raw.startsWith("'")) {
      let i = 1;
      let name = "";
      while (i < raw.length) {
        if (raw[i] === "'" && raw[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        if (raw[i] === "'") {
          i += 1;
          break;
        }
        name += raw[i];
        i += 1;
      }
      if (raw[i] !== "!") {
        throw new Error("sourceRange quoted sheet name must be followed by !A1");
      }
      sourceSheetName = name;
      a1Part = raw.slice(i + 1).trim();
    } else {
      const bang = raw.indexOf("!");
      sourceSheetName = raw.slice(0, bang).trim();
      a1Part = raw.slice(bang + 1).trim();
    }
    if (sourceSheetName === "") {
      throw new Error("sourceRange sheet name must be non-empty");
    }
    // 3D: Sheet1:Sheet3!A1 (colon inside sheet qualifier).
    if (sourceSheetName.includes(":")) {
      throw new Error("sourceRange 3D references are not supported");
    }
  }

  const bareA1 = validateBareA1(a1Part, "sourceRange");
  return {
    sourceSheetName,
    bareA1,
    displaySourceRange: formatSheetQualifiedA1(sourceSheetName, bareA1, chartSheetName),
  };
}

/** Same-sheet A1 only: bare range or matching Sheet!A1. */
export function normalizeSameSheetA1Range(
  sheetName: string,
  value: string,
  fieldName: string,
  ownerName: string,
): string {
  const raw = value.trim();
  if (raw === "") throw new Error(`${fieldName} must be non-empty`);
  let bare = raw;
  if (raw.includes("!")) {
    const bang = raw.lastIndexOf("!");
    const sheetPart = raw
      .slice(0, bang)
      .replace(/^'/, "")
      .replace(/'$/, "")
      .replace(/''/g, "'");
    bare = raw.slice(bang + 1).trim();
    if (sheetPart.toLowerCase() !== sheetName.toLowerCase()) {
      throw new Error(`${fieldName} must be on the same worksheet as the ${ownerName}`);
    }
  }
  return validateBareA1(bare, fieldName);
}

export function normalizeSameSheetSourceRange(sheetName: string, sourceRange: string): string {
  return normalizeSameSheetA1Range(sheetName, sourceRange, "sourceRange", "chart");
}

function toSeriesInfo(item: ExcelChartSeries, index: number): ChartSeriesInfo {
  return {
    index,
    name: requireLoadedString(item.name, "ChartSeries.name"),
    chartType: toChartTypeLabel(requireLoadedString(item.chartType, "ChartSeries.chartType")),
    smooth: requireLoadedBoolean(item.smooth, "ChartSeries.smooth"),
  };
}

/** Replace chart data source via Chart.setData; return real series snapshot. */
export async function officeJsUpdateChartSource(
  input: ChartSourceUpdateInput,
): Promise<HostResult<ChartSourceInfo>> {
  return withExcel("chart.source.update", async (context) => {
    const seriesBy: ChartSeriesBy = input.seriesBy ?? "auto";
    const parsed = parseChartSourceRange(input.sheetName, input.sourceRange);

    const chartSheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = chartSheet.charts.getItem(input.chartName) as unknown as ExcelChartWithSetData;
    if (typeof chart.setData !== "function") {
      throw new Error("Chart.setData is not available in this host");
    }

    // Range proxy from source sheet (may differ from chart sheet); same-workbook only.
    const sourceSheet = context.workbook.worksheets.getItem(parsed.sourceSheetName);
    const range = sourceSheet.getRange(parsed.bareA1);
    chart.setData(range, SERIES_BY_OFFICE[seriesBy]);
    await context.sync();

    chart.series.load("items/name,items/chartType,items/smooth");
    chart.load("name");
    await context.sync();

    const series = chart.series.items.map((item, i) => toSeriesInfo(item, i + 1));
    return {
      sheetName: input.sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      sourceRange: parsed.displaySourceRange,
      seriesBy,
      series,
    };
  });
}
