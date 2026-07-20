/**
 * WPS JSA: no verified workbook template surface — typed unsupported for apply/capture.
 * Do not partially succeed via range.format; COM/.NET/Shell fallback forbidden.
 */
import type { TemplateHostMethods } from "./templateHostMethods";
import type {
  WorkbookTemplateApplyInfo,
  WorkbookTemplateApplyInput,
  WorkbookTemplateCaptureInfo,
} from "./workbookTemplateTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const EVIDENCE =
  "No verified WPS JSA workbook template contract in this repository; COM/.NET/Shell fallback forbidden";

function unsup<T>(capability: string): Promise<HostResult<T>> {
  return Promise.resolve(
    unsupported(
      capability,
      "wps-jsa",
      "Workbook template tools are not verified for WPS JSA in this repository",
      EVIDENCE,
    ) as HostResult<T>,
  );
}

export async function wpsApplyWorkbookTemplate(
  _input: WorkbookTemplateApplyInput,
): Promise<HostResult<WorkbookTemplateApplyInfo>> {
  return unsup("workbook.template.apply");
}

export async function wpsCaptureWorkbookTemplate(): Promise<
  HostResult<WorkbookTemplateCaptureInfo>
> {
  return unsup("workbook.template.capture");
}

/** Base class so WpsJsaSlicerSupport/WpsJsaAdapter stay under line budget. */
export abstract class WpsJsaTemplateSupport implements TemplateHostMethods {
  applyWorkbookTemplate = wpsApplyWorkbookTemplate;
  captureWorkbookTemplate = wpsCaptureWorkbookTemplate;
}
