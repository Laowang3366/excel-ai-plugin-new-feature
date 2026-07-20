/**
 * ExcelApi 1.8 precheck for PivotTable hierarchies / collection.add.
 * Fail-safe: missing isSetSupported or throw → typed unsupported.
 */
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const VERSION = "1.8";
const EVIDENCE =
  "Worksheet.pivotTables.add + row/column/filter/data hierarchies require ExcelApi 1.8";

export function isExcelApi18SupportedForPivot(): boolean {
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
    return Boolean(isSetSupported.call(office!.context!.requirements, REQUIREMENT_SET, VERSION));
  } catch {
    return false;
  }
}

export function requireExcelApi18ForPivot(capability: string): HostResult<never> | null {
  if (isExcelApi18SupportedForPivot()) return null;
  return unsupported(
    capability,
    "office-js",
    "ExcelApi 1.8 is not supported in this host (Office.context.requirements.isSetSupported)",
    EVIDENCE,
  ) as HostResult<never>;
}
