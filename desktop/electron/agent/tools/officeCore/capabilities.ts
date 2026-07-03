/**
 * Office action 能力声明
 *
 * 关联模块：
 * - officeActionAdapter.ts: 根据能力决定 Open XML 路由或 needsCom。
 * - prompts/sections/officeToolsPrompt.ts: 提示词描述这些能力。
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
  { app: "word", operation: "applyHeadingStyles", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "insertOrUpdateToc", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "styleTables", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "insertOrReplaceImage", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "setHeaderFooter", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "snapshot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "addSlide", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "addSlides", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "appendSlide", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "appendSlides", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "addSlideContent", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "applyTheme", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "deleteSlides", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "normalizeLayouts", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "insertChart", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "replacePictureSlot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "alignShapes", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "snapshot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
];

export function findOfficeCapability(app: OfficeActionApp, operation: string): OfficeCapability | undefined {
  return OFFICE_CAPABILITIES.find((capability) => capability.app === app && capability.operation === operation);
}
