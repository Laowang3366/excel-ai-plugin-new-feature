import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";

/**
 * Fail-safe: missing isSetSupported or throw → false (typed unsupported).
 * Matches chart dataLabels / advanced ExcelApi precheck pattern.
 */
export function isExcelApiSupported(version: "1.6" | "1.8"): boolean {
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

export function requireExcelApiForCf(
  capability: string,
): HostResult<never> | null {
  if (isExcelApiSupported("1.6")) return null;
  return unsupported(
    capability,
    "office-js",
    "ExcelApi 1.6 is not supported in this host (Office.context.requirements.isSetSupported)",
    "Range.conditionalFormats requires ExcelApi 1.6",
  ) as HostResult<never>;
}

export function requireExcelApiForDv(
  capability: string,
): HostResult<never> | null {
  if (isExcelApiSupported("1.8")) return null;
  return unsupported(
    capability,
    "office-js",
    "ExcelApi 1.8 is not supported in this host (Office.context.requirements.isSetSupported)",
    "Range.dataValidation requires ExcelApi 1.8",
  ) as HostResult<never>;
}
