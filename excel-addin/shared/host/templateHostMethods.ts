/** Workbook template method group for HostAdapter (line-budget isolation). */
import type {
  WorkbookTemplateApplyInfo,
  WorkbookTemplateApplyInput,
  WorkbookTemplateCaptureInfo,
} from "./workbookTemplateTypes";
import type { HostResult } from "./types";

export interface TemplateHostMethods {
  applyWorkbookTemplate(
    input: WorkbookTemplateApplyInput,
  ): Promise<HostResult<WorkbookTemplateApplyInfo>>;
  captureWorkbookTemplate(): Promise<HostResult<WorkbookTemplateCaptureInfo>>;
}
