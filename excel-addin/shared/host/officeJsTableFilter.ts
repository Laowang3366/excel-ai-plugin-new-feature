/**
 * Office.js Table.autoFilter: get (1.9 enabled), apply/clear (1.2).
 * Public columnIndex is 1-based; Office.js columnIndex is 0-based.
 */
import { withExcel } from "./officeJsRuntime";
import type {
  TableFilterApplyInput,
  TableFilterClearInput,
  TableFilterGetInput,
  TableFilterInfo,
  TableFilterOn,
} from "./tableFilterTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const APPLY_EVIDENCE =
  "Table.autoFilter.apply / clearCriteria require ExcelApi 1.2";
const GET_EVIDENCE =
  "Table.autoFilter.enabled requires ExcelApi 1.9";

type ExcelApiVersion = "1.2" | "1.9";

function isExcelApiSupported(version: ExcelApiVersion): boolean {
  const office = (
    globalThis as unknown as {
      Office?: {
        context?: {
          requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
        };
      };
    }
  ).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", version);
  } catch {
    return false;
  }
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded non-empty string`);
  }
  return value;
}

/** Map public filterOn → Office.js FilterOn string enum. */
function toOfficeFilterOn(filterOn: TableFilterOn): string {
  switch (filterOn) {
    case "values":
      return "Values";
    case "custom":
      return "Custom";
    case "topItems":
      return "TopItems";
    case "bottomItems":
      return "BottomItems";
    case "topPercent":
      return "TopPercent";
    case "bottomPercent":
      return "BottomPercent";
    default: {
      const _exhaustive: never = filterOn;
      return _exhaustive;
    }
  }
}

function buildCriteria(input: TableFilterApplyInput): Record<string, unknown> {
  const filterOn = toOfficeFilterOn(input.filterOn);
  switch (input.filterOn) {
    case "values": {
      if (!input.values || input.values.length === 0) {
        throw new Error("filterOn=values requires non-empty values[]");
      }
      return { filterOn, values: input.values };
    }
    case "custom": {
      if (input.criterion1 == null || String(input.criterion1).trim() === "") {
        throw new Error("filterOn=custom requires criterion1");
      }
      const criteria: Record<string, unknown> = {
        filterOn,
        criterion1: String(input.criterion1),
      };
      if (input.criterion2 != null && String(input.criterion2).trim() !== "") {
        criteria.criterion2 = String(input.criterion2);
      }
      if (input.operator != null) {
        criteria.operator = input.operator === "or" ? "Or" : "And";
      }
      return criteria;
    }
    case "topItems":
    case "bottomItems":
    case "topPercent":
    case "bottomPercent": {
      if (input.threshold == null || !Number.isFinite(input.threshold) || input.threshold <= 0) {
        throw new Error(`${input.filterOn} requires positive threshold`);
      }
      // Office.js uses criterion1 as the numeric threshold string for top/bottom.
      return { filterOn, criterion1: String(input.threshold) };
    }
    default: {
      const _exhaustive: never = input.filterOn;
      return _exhaustive;
    }
  }
}

export async function officeJsGetTableFilter(
  input: TableFilterGetInput,
): Promise<HostResult<TableFilterInfo>> {
  if (!isExcelApiSupported("1.9")) {
    return unsupported(
      "table.filter.get",
      "office-js",
      "ExcelApi 1.9 is not supported in this host (Office.context.requirements.isSetSupported)",
      GET_EVIDENCE,
    );
  }
  return withExcel("table.filter.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    sheet.load("name");
    table.load("name");
    const autoFilter = table.autoFilter;
    if (autoFilter == null || typeof autoFilter.load !== "function") {
      throw new Error("Table.autoFilter missing (ExcelApi 1.2+ required)");
    }
    autoFilter.load("enabled");
    await context.sync();
    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      tableName: requireLoadedString(table.name, "Table.name"),
      enabled: Boolean(autoFilter.enabled),
    };
  });
}

export async function officeJsApplyTableFilter(
  input: TableFilterApplyInput,
): Promise<HostResult<TableFilterInfo>> {
  if (!isExcelApiSupported("1.2")) {
    return unsupported(
      "table.filter.apply",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      APPLY_EVIDENCE,
    );
  }
  if (!Number.isInteger(input.columnIndex) || input.columnIndex < 1) {
    throw new Error("columnIndex must be a 1-based integer >= 1");
  }
  const criteria = buildCriteria(input);
  const zeroBased = input.columnIndex - 1;

  return withExcel("table.filter.apply", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    sheet.load("name");
    table.load("name");
    const autoFilter = table.autoFilter;
    if (autoFilter == null || typeof autoFilter.apply !== "function") {
      throw new Error("Table.autoFilter.apply missing (ExcelApi 1.2 required)");
    }
    // Prefer table range so columnIndex is 0-based within the table.
    const range =
      typeof table.getRange === "function" ? table.getRange() : (undefined as never);
    autoFilter.apply(range as never, zeroBased, criteria as never);
    await context.sync();

    let enabled = true;
    if (isExcelApiSupported("1.9") && typeof autoFilter.load === "function") {
      autoFilter.load("enabled");
      await context.sync();
      enabled = Boolean(autoFilter.enabled);
    }

    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      tableName: requireLoadedString(table.name, "Table.name"),
      enabled,
      columnIndex: input.columnIndex,
      filterOn: input.filterOn,
    };
  });
}

export async function officeJsClearTableFilter(
  input: TableFilterClearInput,
): Promise<HostResult<TableFilterInfo>> {
  if (!isExcelApiSupported("1.2")) {
    return unsupported(
      "table.filter.clear",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      APPLY_EVIDENCE,
    );
  }
  return withExcel("table.filter.clear", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    sheet.load("name");
    table.load("name");
    const autoFilter = table.autoFilter;
    if (autoFilter == null || typeof autoFilter.clearCriteria !== "function") {
      throw new Error("Table.autoFilter.clearCriteria missing (ExcelApi 1.2 required)");
    }
    autoFilter.clearCriteria();
    await context.sync();

    let enabled = false;
    if (isExcelApiSupported("1.9") && typeof autoFilter.load === "function") {
      autoFilter.load("enabled");
      await context.sync();
      enabled = Boolean(autoFilter.enabled);
    }

    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      tableName: requireLoadedString(table.name, "Table.name"),
      enabled,
    };
  });
}
