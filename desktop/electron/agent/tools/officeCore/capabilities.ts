/**
 * Office action 能力声明
 *
 * 关联模块：
 * - officeActionAdapter.ts: 根据能力决定 Open XML 路由或 needsCom。
 * - prompts/templates/scenarios/office-tools.zh-CN.md: 提示词描述这些能力。
 */

import type { OfficeActionApp, OfficeActionEngine } from "./types";

export interface OfficeCapability {
  app: OfficeActionApp;
  operation: string;
  preferredEngine: OfficeActionEngine;
  writesFile: boolean;
  fallback: "none" | "needsCom";
}

export const OFFICE_CAPABILITIES: OfficeCapability[] = [
  { app: "excel", operation: "insertChart", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "applyConditionalFormatting", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "setDataValidation", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "styleTable", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "excel", operation: "snapshot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "createPivotTable", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "refreshPivotTables", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "addSlicer", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "createPowerQuery", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectPowerQueries", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "managePowerQuery", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectCharts", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "formatChart", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectWorkbookObjects", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "manageWorkbookObject", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "manageWorksheetObjects", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "captureWorkbookTemplate", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "inspectWorkbookFormatting", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "applyWorkbookTemplate", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectPrintSettings", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "configurePrint", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "exportSheetsToPdf", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "exportPdf", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "traceFormulaDependencies", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "inspectFormulaDependencies", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "repairFormulaReferences", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "convertFormulasToValues", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectFormulaBackups", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "restoreFormulas", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "inspectFormulaProtection", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "excel", operation: "manageFormulaProtection", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "exportRangeToWord", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "exportRangeToPresentation", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "buildReportPackage", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "applyHeadingStyles", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "insertOrUpdateToc", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "styleTables", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "insertOrReplaceImage", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "setHeaderFooter", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "snapshot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "inspectDocumentFormatting", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "word", operation: "formatLongDocument", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "inspectReferences", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "word", operation: "manageReferences", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "inspectRevisions", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "word", operation: "manageRevisions", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "compareDocuments", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "applyTrackedChanges", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "prepareMailMergeTemplate", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "mailMerge", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "batchMailMerge", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "inspectContentControls", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "word", operation: "populateContentControls", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "manageContentControls", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "exportPdf", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "inspectLinkedOfficeContent", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "word", operation: "refreshLinkedOfficeContent", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "relinkLinkedOfficeContent", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "addSlide", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "addSlides", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "appendSlide", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "appendSlides", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "addSlideContent", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "applyTheme", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "deleteSlides", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "normalizeLayouts", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "insertChart", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "insertTable", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "replacePictureSlot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "alignShapes", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "snapshot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "inspectPresentationTheme", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "presentation", operation: "inspectSlideElements", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "presentation", operation: "inspectAnimations", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "presentation", operation: "inspectSpeakerNotes", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "presentation", operation: "applyMasterBranding", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "layoutElements", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "configureAnimations", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "configureSlideShow", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "setSpeakerNotes", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "exportHandouts", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "inspectLinkedOfficeContent", preferredEngine: "com", writesFile: false, fallback: "needsCom" },
  { app: "presentation", operation: "refreshLinkedOfficeContent", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "relinkLinkedOfficeContent", preferredEngine: "com", writesFile: true, fallback: "needsCom" },
];

export function findOfficeCapability(app: OfficeActionApp, operation: string): OfficeCapability | undefined {
  return OFFICE_CAPABILITIES.find((capability) => capability.app === app && capability.operation === operation);
}
