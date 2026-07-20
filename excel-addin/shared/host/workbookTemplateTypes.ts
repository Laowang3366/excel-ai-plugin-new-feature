/**
 * Desktop parity: applyWorkbookTemplate + captureWorkbookTemplate/inspectWorkbookFormatting.
 * Capture is one safe tool; do not add a synonym inspect tool.
 */

export const WORKBOOK_TEMPLATE_PRESETS = [
  "professional",
  "financial",
  "dashboard",
  "minimal",
] as const;

export type WorkbookTemplatePreset = (typeof WORKBOOK_TEMPLATE_PRESETS)[number];

export function isWorkbookTemplatePreset(value: unknown): value is WorkbookTemplatePreset {
  return typeof value === "string" && (WORKBOOK_TEMPLATE_PRESETS as readonly string[]).includes(value);
}

export interface WorkbookTemplatePresetStyle {
  headerFill: string;
  headerFontColor: string;
}

/** Desktop Worker colors (hex without #). */
export const WORKBOOK_TEMPLATE_PRESET_STYLES: Record<
  WorkbookTemplatePreset,
  WorkbookTemplatePresetStyle
> = {
  professional: { headerFill: "#1F4E79", headerFontColor: "#FFFFFF" },
  financial: { headerFill: "#217346", headerFontColor: "#FFFFFF" },
  dashboard: { headerFill: "#202124", headerFontColor: "#FFFFFF" },
  minimal: { headerFill: "#E8EAED", headerFontColor: "#202124" },
};

export interface WorkbookTemplateApplyInput {
  preset: WorkbookTemplatePreset;
  sheetNames?: string[];
  allSheets: boolean;
  fontName: string;
  fontSize: number;
  autoFit: boolean;
  showGridlines: boolean;
  freezeRows: number;
}

export interface WorkbookTemplateSheetReadback {
  fontName: string;
  fontSize: number;
  headerFill: string;
  headerFontColor: string;
  headerBold: boolean;
  headerHorizontalAlignment: string;
  headerWrapText: boolean;
  headerRowHeight: number;
  showGridlines: boolean;
  freezeRowCount: number;
  autoFitVerified: false;
}

export interface WorkbookTemplateAppliedSheet {
  name: string;
  range: string;
  rows: number;
  columns: number;
  readback: WorkbookTemplateSheetReadback;
}

export interface WorkbookTemplateSkippedSheet {
  name: string;
  reason: string;
}

export interface WorkbookTemplateApplyInfo {
  preset: WorkbookTemplatePreset;
  appliedSheets: WorkbookTemplateAppliedSheet[];
  appliedSheetCount: number;
  skippedSheets: WorkbookTemplateSkippedSheet[];
  limitations: string[];
}

export interface WorkbookTemplateBaseStyle {
  fontName: string | null;
  fontSize: number | null;
  fontColor: string | null;
}

export interface WorkbookTemplateHeaderStyle {
  fillColor: string | null;
  fontColor: string | null;
  bold: boolean | null;
  rowHeight: number | null;
}

export interface WorkbookTemplatePrintSnapshot {
  area: string | null;
  orientation: string | null;
  paperSize: string | null;
  fitToPagesWide: number | null;
  fitToPagesTall: number | null;
  repeatRows: string | null;
  repeatColumns: string | null;
  header: string | null;
  footer: string | null;
}

export interface WorkbookTemplateCapturedSheet {
  name: string;
  usedRange: string | null;
  rows: number;
  columns: number;
  baseStyle: WorkbookTemplateBaseStyle | null;
  headerStyle: WorkbookTemplateHeaderStyle | null;
  print: WorkbookTemplatePrintSnapshot;
  limitations: string[];
}

export interface WorkbookTemplateCaptureTemplate {
  version: 1;
  capturedFrom: string;
  capturedAt: string;
  sheets: WorkbookTemplateCapturedSheet[];
}

export interface WorkbookTemplateCaptureInfo {
  template: WorkbookTemplateCaptureTemplate;
  sheetCount: number;
  limitations: string[];
}

export const WORKBOOK_TEMPLATE_MAX_SHEETS = 500;
export const WORKBOOK_TEMPLATE_HEADER_ROW_HEIGHT = 24;
export const WORKBOOK_TEMPLATE_DEFAULT_FONT_NAME = "微软雅黑";
export const WORKBOOK_TEMPLATE_DEFAULT_FONT_SIZE = 10.5;
