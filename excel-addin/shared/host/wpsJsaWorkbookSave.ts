/**
 * workbook.save — WPS JSA ActiveWorkbook.Save via member probe.
 * Evidence: desktop/public/wps-jsa-bridge/main.js uses ActiveWorkbook.Save after CodeModule write.
 * Not an official JSA contract; missing Save → typed unsupported. No saveAs.
 */
import { requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult, WorkbookSaveInfo } from "./types";
import { fail, ok, unsupported } from "./types";

const CAPABILITY = "workbook.save";
const EVIDENCE =
  "desktop/public/wps-jsa-bridge/main.js ActiveWorkbook.Save + typeof Save === 'function' probe";

type WpsWorkbookWithSave = {
  Name: string;
  Save?: unknown;
};

export async function wpsSaveWorkbook(): Promise<HostResult<WorkbookSaveInfo>> {
  const workbookResult = requireWorkbook(CAPABILITY);
  if (!workbookResult.ok) return workbookResult;
  const workbook = workbookResult.data as WpsWorkbookWithSave;
  if (typeof workbook.Save !== "function") {
    return unsupported(
      CAPABILITY,
      "wps-jsa",
      "ActiveWorkbook.Save is not a function",
      EVIDENCE,
    );
  }
  try {
    (workbook.Save as () => void).call(workbook);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(CAPABILITY, "wps-jsa", message, EVIDENCE);
  }
  const name = workbook.Name;
  if (typeof name !== "string" || name.trim() === "") {
    return fail(
      CAPABILITY,
      "wps-jsa",
      "ActiveWorkbook.Name is empty after Save",
      EVIDENCE,
    );
  }
  return ok({ workbookName: name, saved: true });
}
