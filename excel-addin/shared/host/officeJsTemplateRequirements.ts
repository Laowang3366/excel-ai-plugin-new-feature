/**
 * Workbook template requirement prechecks.
 * apply: ExcelApi 1.8 (showGridlines + freeze 1.7 + format 1.1 + autofit 1.2 subset under 1.8).
 * capture: ExcelApi 1.9 (pageLayout print snapshot).
 */
import type { HostResult } from "./types";
import { unsupported } from "./types";

const APPLY_VERSION = "1.8";
const CAPTURE_VERSION = "1.9";

const APPLY_EVIDENCE =
  "Workbook template apply needs ExcelApi 1.8 (Worksheet.showGridlines 1.8; freezePanes 1.7; Range.format/autofit ≤1.2)";
const CAPTURE_EVIDENCE =
  "Workbook template capture needs ExcelApi 1.9 (PageLayout print snapshot + UsedRange format read)";

function isSetSupported(version: string): boolean {
  const office = (
    globalThis as unknown as {
      Office?: {
        context?: {
          requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
        };
      };
    }
  ).Office;
  const fn = office?.context?.requirements?.isSetSupported;
  if (typeof fn !== "function") return false;
  try {
    return Boolean(fn.call(office!.context!.requirements, "ExcelApi", version));
  } catch {
    return false;
  }
}

export function isExcelApi18ForTemplateApply(): boolean {
  return isSetSupported(APPLY_VERSION);
}

export function isExcelApi19ForTemplateCapture(): boolean {
  return isSetSupported(CAPTURE_VERSION);
}

export function requireExcelApi18ForTemplateApply(
  capability: string,
): HostResult<never> | null {
  if (isExcelApi18ForTemplateApply()) return null;
  return unsupported(
    capability,
    "office-js",
    `ExcelApi ${APPLY_VERSION} is not supported in this host (Office.context.requirements.isSetSupported)`,
    APPLY_EVIDENCE,
  ) as HostResult<never>;
}

export function requireExcelApi19ForTemplateCapture(
  capability: string,
): HostResult<never> | null {
  if (isExcelApi19ForTemplateCapture()) return null;
  return unsupported(
    capability,
    "office-js",
    `ExcelApi ${CAPTURE_VERSION} is not supported in this host (Office.context.requirements.isSetSupported)`,
    CAPTURE_EVIDENCE,
  ) as HostResult<never>;
}
