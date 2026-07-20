/** Office.js workbook template method group for OfficeJsAdapter binding. */
import { officeJsApplyWorkbookTemplate } from "./officeJsTemplateApply";
import { officeJsCaptureWorkbookTemplate } from "./officeJsTemplateCapture";
import type { TemplateHostMethods } from "./templateHostMethods";

export const officeJsTemplateMethods: TemplateHostMethods = {
  applyWorkbookTemplate: officeJsApplyWorkbookTemplate,
  captureWorkbookTemplate: officeJsCaptureWorkbookTemplate,
};
