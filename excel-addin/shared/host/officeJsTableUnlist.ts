import { withExcel } from "./officeJsRuntime";
import type { HostResult, TableUnlistInfo } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_EVIDENCE =
  "Table.convertToRange requires ExcelApi 1.2 to convert a table to a plain range while keeping data";

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded non-empty string`);
  }
  return value;
}

/** Official precheck before any convertToRange / workbook access for table.unlist. */
export function isExcelApi12SupportedForTableUnlist(): boolean {
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

/**
 * Convert table to range (keep data). Flow:
 * precheck 1.2 → load sheet/table/address → sync → convertToRange → sync →
 * load table names → sync → case-insensitive absence check → host result.
 */
export async function officeJsUnlistTable(
  sheetName: string,
  tableName: string,
): Promise<HostResult<TableUnlistInfo>> {
  if (!isExcelApi12SupportedForTableUnlist()) {
    return unsupported(
      "table.unlist",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  const result = await withExcel("table.unlist", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItem(tableName);
    if (typeof table.convertToRange !== "function") {
      throw new Error("Table.convertToRange missing (ExcelApi 1.2 required)");
    }
    sheet.load("name");
    table.load("name");
    const range = table.getRange();
    range.load("address");
    await context.sync();

    const hostSheetName = requireLoadedString(sheet.name, "Worksheet.name");
    const hostTableName = requireLoadedString(table.name, "Table.name");
    const hostAddress = requireLoadedString(range.address, "Range.address");

    table.convertToRange();
    await context.sync();

    sheet.tables.load("items/name");
    await context.sync();

    const stillPresent = sheet.tables.items.some(
      (item) =>
        typeof item.name === "string" &&
        item.name.toLowerCase() === hostTableName.toLowerCase(),
    );
    if (stillPresent) {
      throw new Error(
        `Table "${hostTableName}" still present after convertToRange absence check`,
      );
    }

    return {
      sheetName: hostSheetName,
      tableName: hostTableName,
      address: hostAddress,
      unlisted: true as const,
    };
  });

  return result;
}
