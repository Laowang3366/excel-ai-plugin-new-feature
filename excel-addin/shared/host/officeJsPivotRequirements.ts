/**
 * Pivot requirement-set prechecks.
 * - list/create + hierarchy layout: ExcelApi 1.8
 * - refresh (PivotTable.refresh / collection items name): ExcelApi 1.3
 * Fail-safe: missing isSetSupported or throw → typed unsupported.
 */
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";

export type PivotExcelApiVersion = "1.3" | "1.8";

const EVIDENCE: Record<PivotExcelApiVersion, string> = {
  "1.3": "PivotTable.refresh / PivotTableCollection items require ExcelApi 1.3",
  "1.8":
    "Worksheet.pivotTables.add + row/column/filter/data hierarchies require ExcelApi 1.8",
};

export function isExcelApiSupportedForPivot(version: PivotExcelApiVersion): boolean {
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
    return Boolean(isSetSupported.call(office!.context!.requirements, REQUIREMENT_SET, version));
  } catch {
    return false;
  }
}

/** @deprecated prefer isExcelApiSupportedForPivot("1.8") */
export function isExcelApi18SupportedForPivot(): boolean {
  return isExcelApiSupportedForPivot("1.8");
}

export function requireExcelApiForPivot(
  capability: string,
  version: PivotExcelApiVersion,
): HostResult<never> | null {
  if (isExcelApiSupportedForPivot(version)) return null;
  return unsupported(
    capability,
    "office-js",
    `ExcelApi ${version} is not supported in this host (Office.context.requirements.isSetSupported)`,
    EVIDENCE[version],
  ) as HostResult<never>;
}

export function requireExcelApi18ForPivot(capability: string): HostResult<never> | null {
  return requireExcelApiForPivot(capability, "1.8");
}

export function requireExcelApi13ForPivotRefresh(capability: string): HostResult<never> | null {
  return requireExcelApiForPivot(capability, "1.3");
}
