/**
 * Office.js Table.sort: apply / clear / fields readback (ExcelApi 1.2).
 * Public columnIndex is 1-based; SortField.key is 0-based.
 */
import { withExcel } from "./officeJsRuntime";
import type {
  TableSortApplyInput,
  TableSortClearInput,
  TableSortFieldInfo,
  TableSortGetInput,
  TableSortInfo,
} from "./tableSortTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const SORT_EVIDENCE = "Table.sort.apply / clear / fields require ExcelApi 1.2";

function isExcelApi12Supported(): boolean {
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
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.2");
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

function toOfficeSortFields(input: TableSortApplyInput): Array<{ key: number; ascending: boolean }> {
  if (!Array.isArray(input.fields) || input.fields.length === 0) {
    throw new Error("fields must be a non-empty array");
  }
  if (input.fields.length > 3) {
    throw new Error("fields supports at most 3 levels");
  }
  return input.fields.map((field, index) => {
    if (!Number.isInteger(field.columnIndex) || field.columnIndex < 1) {
      throw new Error(`fields[${index}].columnIndex must be a 1-based integer >= 1`);
    }
    return {
      key: field.columnIndex - 1,
      ascending: field.ascending !== false,
    };
  });
}

async function readSortInfo(
  context: { sync: () => Promise<void> },
  sheet: { name: string; load: (p: string) => void },
  table: {
    name: string;
    load: (p: string) => void;
    sort: {
      fields: {
        load: (p: string) => void;
        items?: Array<{ key?: number; ascending?: boolean; load?: (p: string) => void }>;
      };
    };
  },
): Promise<TableSortInfo> {
  sheet.load("name");
  table.load("name");
  const fields = table.sort.fields;
  if (fields == null || typeof fields.load !== "function") {
    throw new Error("Table.sort.fields missing (ExcelApi 1.2 required)");
  }
  fields.load("items");
  await context.sync();

  const items = fields.items ?? [];
  for (const item of items) {
    if (item && typeof item.load === "function") {
      item.load("key,ascending");
    }
  }
  if (items.length > 0) {
    await context.sync();
  }

  const mapped: TableSortFieldInfo[] = items.map((item, index) => {
    const key = item?.key;
    if (typeof key !== "number" || !Number.isFinite(key) || key < 0) {
      throw new Error(`sort.fields[${index}].key is not a loaded non-negative number`);
    }
    return {
      columnIndex: key + 1,
      ascending: item.ascending !== false,
    };
  });

  return {
    sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
    tableName: requireLoadedString(table.name, "Table.name"),
    fields: mapped,
  };
}

export async function officeJsGetTableSort(
  input: TableSortGetInput,
): Promise<HostResult<TableSortInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "table.sort.get",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      SORT_EVIDENCE,
    );
  }
  return withExcel("table.sort.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    if (table.sort == null) {
      throw new Error("Table.sort missing (ExcelApi 1.2 required)");
    }
    return readSortInfo(context, sheet, table);
  });
}

export async function officeJsApplyTableSort(
  input: TableSortApplyInput,
): Promise<HostResult<TableSortInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "table.sort.apply",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      SORT_EVIDENCE,
    );
  }
  const officeFields = toOfficeSortFields(input);
  return withExcel("table.sort.apply", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    const sort = table.sort;
    if (sort == null || typeof sort.apply !== "function") {
      throw new Error("Table.sort.apply missing (ExcelApi 1.2 required)");
    }
    sort.apply(officeFields as never, input.matchCase === true);
    await context.sync();
    return readSortInfo(context, sheet, table);
  });
}

export async function officeJsClearTableSort(
  input: TableSortClearInput,
): Promise<HostResult<TableSortInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "table.sort.clear",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      SORT_EVIDENCE,
    );
  }
  return withExcel("table.sort.clear", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    const sort = table.sort;
    if (sort == null || typeof sort.clear !== "function") {
      throw new Error("Table.sort.clear missing (ExcelApi 1.2 required)");
    }
    sort.clear();
    await context.sync();
    return readSortInfo(context, sheet, table);
  });
}
