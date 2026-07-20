/**
 * List DataValidation source materialization (inline string vs Range ClientObject).
 */
import { normalizeSameSheetA1Range, parseChartSourceRange } from "./officeJsChartSource";
import type { ExcelRange, ExcelRequestContext } from "./officeJsRuntime";
import { classifyListSource } from "./officeJsValidationCompare";

export type ExcelClientObjectRange = {
  load: (props: string) => void;
  address: string;
};

export type ExcelPlainAddressRange = {
  address: string;
};

/** Real Office.js Range proxy: has load(); do NOT read address before load+sync. */
export function isClientObjectRange(value: unknown): value is ExcelClientObjectRange {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { load?: unknown }).load === "function"
  );
}

/** Already-materialized plain { address } (tests / non-ClientObject). */
export function isPlainAddressRange(value: unknown): value is ExcelPlainAddressRange {
  if (typeof value !== "object" || value === null) return false;
  if (typeof (value as { load?: unknown }).load === "function") return false;
  return typeof (value as { address?: unknown }).address === "string";
}

export function resolveListSourceRange(
  context: ExcelRequestContext,
  ownerSheetName: string,
  formula1: string,
): ExcelRange {
  const raw = formula1.trim().replace(/^=/, "");
  if (!raw.includes("!")) {
    const bare = normalizeSameSheetA1Range(ownerSheetName, raw, "formula1", "dataValidation");
    return context.workbook.worksheets.getItem(ownerSheetName).getRange(bare);
  }
  const parsed = parseChartSourceRange(ownerSheetName, raw);
  return context.workbook.worksheets
    .getItem(parsed.sourceSheetName)
    .getRange(parsed.bareA1);
}


/**
 * Resolve list.source that may be string or Excel.Range ClientObject.
 * When Range-like, load address if needed (caller must sync when load queued).
 */
export async function materializeListSource(
  source: unknown,
  context: ExcelRequestContext,
): Promise<{
  kind: "inline" | "range" | null;
  formula1?: string;
  listValues?: string[];
  lossy?: boolean;
  limitations?: string[];
}> {
  if (source == null) {
    return {
      kind: null,
      limitations: ["list source is empty"],
    };
  }
  if (isClientObjectRange(source)) {
    // Always load+sync before reading address (PropertyNotLoaded otherwise).
    source.load("address");
    await context.sync();
    let address = "";
    try {
      address = String(source.address ?? "").trim();
    } catch (err) {
      return {
        kind: null,
        limitations: [
          `list Range source address PropertyNotLoaded after load: ${
            err instanceof Error ? err.message : String(err)
          }`,
        ],
      };
    }
    if (!address) {
      return {
        kind: null,
        limitations: ["list Range source address unavailable after load"],
      };
    }
    return { kind: "range", formula1: address };
  }
  if (isPlainAddressRange(source)) {
    const address = source.address.trim();
    if (!address) {
      return {
        kind: null,
        limitations: ["list Range source address empty"],
      };
    }
    return { kind: "range", formula1: address };
  }
  if (typeof source !== "string") {
    return {
      kind: null,
      limitations: [
        `list source is not string or Range (got ${typeof source}); not coerced via String(object)`,
      ],
    };
  }
  const classified = classifyListSource(source);
  if (classified.kind == null || classified.lossy) {
    return {
      kind: null,
      formula1: classified.formula1,
      listValues: classified.listValues,
      lossy: true,
      limitations:
        classified.limitations ??
        [`list source is lossy/unparseable: ${classified.raw ?? source}`],
    };
  }
  if (classified.kind === "range") {
    return { kind: "range", formula1: classified.formula1 };
  }
  return { kind: "inline", listValues: classified.listValues };
}
