/**
 * ExcelApi 1.9 Worksheet manual page breaks (not PageLayout).
 * Address readback via PageBreak.getCellAfterBreak().address after load+sync.
 */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelPageBreakCollection, ExcelPageBreak } from "./officeJsPageBreakTypes";
import type { ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";

/** Normalize host Range.address (Sheet1!$A$4 / $C$1 / A4) → bare single-cell A1. */
export function normalizePageBreakAddress(address: unknown): string {
  if (typeof address !== "string" || address.trim() === "") {
    throw new Error("PageBreak cell address is not a loaded non-empty string");
  }
  const bare = (address.includes("!") ? address.split("!")[1]! : address).replace(/\$/g, "").trim();
  // single cell only
  if (bare.includes(":") || bare.includes(",")) {
    throw new Error(`PageBreak cell address is not a single cell A1: ${bare}`);
  }
  const parsed = parseA1Cell(bare);
  if (!parsed) {
    throw new Error(`PageBreak cell address is not a single cell A1: ${bare}`);
  }
  return toA1(parsed.row, parsed.col);
}

function requireCollection(
  sheet: ExcelWorksheet,
  key: "horizontalPageBreaks" | "verticalPageBreaks",
): ExcelPageBreakCollection {
  if (!(key in sheet) || sheet[key] == null || typeof sheet[key] !== "object") {
    throw new Error(`Worksheet.${key} is missing on host sheet object`);
  }
  return sheet[key];
}

async function readCollectionAddresses(
  collection: ExcelPageBreakCollection,
  context: ExcelRequestContext,
  collectionName: string,
): Promise<string[]> {
  if (!("items" in collection)) {
    throw new Error(`Worksheet.${collectionName}.items is missing on host collection`);
  }
  collection.load("items");
  await context.sync();
  const items = collection.items ?? [];
  const cells: Array<{ break: ExcelPageBreak; cell: { address: string; load: (p: string) => void } }> =
    [];
  for (const item of items) {
    if (typeof item.getCellAfterBreak !== "function") {
      throw new Error(
        `Worksheet.${collectionName} item.getCellAfterBreak is missing on host page break`,
      );
    }
    const cell = item.getCellAfterBreak();
    cell.load("address");
    cells.push({ break: item, cell });
  }
  await context.sync();
  return cells.map((entry) => normalizePageBreakAddress(entry.cell.address));
}

export async function readManualPageBreaks(
  sheet: ExcelWorksheet,
  context: ExcelRequestContext,
): Promise<{ horizontalPageBreaks: string[]; verticalPageBreaks: string[] }> {
  const horizontal = requireCollection(sheet, "horizontalPageBreaks");
  const vertical = requireCollection(sheet, "verticalPageBreaks");
  // Read sequentially so each collection's items load is isolated and ordered.
  const horizontalPageBreaks = await readCollectionAddresses(
    horizontal,
    context,
    "horizontalPageBreaks",
  );
  const verticalPageBreaks = await readCollectionAddresses(vertical, context, "verticalPageBreaks");
  return { horizontalPageBreaks, verticalPageBreaks };
}

export function applyManualPageBreaks(
  sheet: ExcelWorksheet,
  input: {
    clearPageBreaks?: boolean;
    horizontalPageBreaks?: string[];
    verticalPageBreaks?: string[];
  },
): void {
  const horizontal = requireCollection(sheet, "horizontalPageBreaks");
  const vertical = requireCollection(sheet, "verticalPageBreaks");

  if (input.clearPageBreaks === true) {
    if (typeof horizontal.removePageBreaks !== "function") {
      throw new Error("Worksheet.horizontalPageBreaks.removePageBreaks is missing on host collection");
    }
    if (typeof vertical.removePageBreaks !== "function") {
      throw new Error("Worksheet.verticalPageBreaks.removePageBreaks is missing on host collection");
    }
    horizontal.removePageBreaks();
    vertical.removePageBreaks();
  }

  if (input.horizontalPageBreaks) {
    if (typeof horizontal.add !== "function") {
      throw new Error("Worksheet.horizontalPageBreaks.add is missing on host collection");
    }
    for (const address of input.horizontalPageBreaks) {
      horizontal.add(address);
    }
  }
  if (input.verticalPageBreaks) {
    if (typeof vertical.add !== "function") {
      throw new Error("Worksheet.verticalPageBreaks.add is missing on host collection");
    }
    for (const address of input.verticalPageBreaks) {
      vertical.add(address);
    }
  }
}
