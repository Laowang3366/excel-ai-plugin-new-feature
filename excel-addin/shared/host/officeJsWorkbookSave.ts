/**
 * workbook.save — Office.js Excel.Workbook.save (ExcelApi 1.1).
 * No saveAs: ExcelApi 1.8 saveAs is online-only and path-based save is out of scope.
 */
import { withExcel } from "./officeJsRuntime";
import type { HostResult, WorkbookSaveInfo } from "./types";
import { unsupported } from "./types";

const CAPABILITY = "workbook.save";

/** Fail-safe: missing isSetSupported or throw → unsupported. */
export function isExcelApi11SupportedForSave(): boolean {
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
    return Boolean(isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.1"));
  } catch {
    return false;
  }
}

export async function officeJsSaveWorkbook(): Promise<HostResult<WorkbookSaveInfo>> {
  if (!isExcelApi11SupportedForSave()) {
    return unsupported(
      CAPABILITY,
      "office-js",
      "ExcelApi 1.1 is not supported in this host (Office.context.requirements.isSetSupported)",
      "Excel.Workbook.save requires ExcelApi 1.1; saveAs/path not implemented",
    );
  }
  return withExcel(CAPABILITY, async (context) => {
    context.workbook.load("name");
    context.workbook.save();
    await context.sync();
    // Post-sync name readback (already loaded or re-load if host cleared).
    const name = context.workbook.name;
    if (typeof name !== "string" || name.trim() === "") {
      throw new Error("Workbook.name is not a loaded non-empty string after save");
    }
    return { workbookName: name, saved: true as const };
  });
}
