/**
 * Slicer requirement-set precheck — ExcelApi 1.10 (official stable).
 * Fail-safe: missing isSetSupported or throw → typed unsupported.
 */
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const VERSION = "1.10";
const EVIDENCE =
  "Excel.Slicer / SlicerCollection / SlicerItem require ExcelApi 1.10 (stable; not BETA nameInFormula/slicerStyle/setStyle)";

export function isExcelApi110SupportedForSlicer(): boolean {
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

export function requireExcelApi110ForSlicer(capability: string): HostResult<never> | null {
  if (isExcelApi110SupportedForSlicer()) return null;
  return unsupported(
    capability,
    "office-js",
    `ExcelApi ${VERSION} is not supported in this host (Office.context.requirements.isSetSupported)`,
    EVIDENCE,
  ) as HostResult<never>;
}
