import path from "path";
import { psVar } from "../../../automation/powershell";
import { parseOfficeLocator } from "../../officeCore/locator";
import type { OfficeActionInput } from "../../officeCore/types";
export function stringParam(input: OfficeActionInput, key: string): string | undefined {
  return typeof input.params?.[key] === "string" ? input.params[key] : undefined;
}

export function actionParamsScript(input: OfficeActionInput): string {
  return `${psVar("_paramsJson", JSON.stringify(input.params || {}))}\n$actionParams = ConvertFrom-Json $_paramsJson`;
}

export function outputPathForAction(input: OfficeActionInput, exportSuffix?: string): string {
  if (input.outputPath) return input.outputPath;
  if (exportSuffix) return defaultOutputPath(input.filePath!, exportSuffix);
  return input.filePath!;
}

export function validationFormulaParam(input: OfficeActionInput): string {
  const values = input.params?.values;
  if (Array.isArray(values)) return values.map((item) => String(item)).join(",");
  return stringParam(input, "formula") || stringParam(input, "values") || "";
}

export function targetSlideIndex(target?: string): number {
  const locator = target ? parseOfficeLocator(target) : undefined;
  return locator?.kind === "slide" && locator.index ? locator.index : 1;
}

export function presentationDeleteSlideIndexesLiteral(input: OfficeActionInput): string {
  const slides = Array.isArray(input.params?.slides)
    ? input.params.slides.map((item) => Number(item)).filter((item) => Number.isFinite(item))
    : [];
  const range = slides.length > 0 ? slides : presentationDeleteRange(input);
  const indexes = [...new Set(range.map((item) => Math.floor(item)).filter((item) => item >= 1))].sort((a, b) => a - b);
  return indexes.length > 0 ? `@(${indexes.join(", ")})` : "@()";
}

export function presentationDeleteRange(input: OfficeActionInput): number[] {
  const from = numericParam(input, "from") ?? numericParam(input, "start");
  const to = numericParam(input, "to") ?? numericParam(input, "end") ?? from;
  if (from && to) return buildNumberRange(from, to);
  const target = input.target || "";
  const match = target.match(/^slides?:\s*(\d+)(?:\s*-\s*(\d+))?$/i);
  if (!match) return [];
  return buildNumberRange(Number(match[1]), Number(match[2] || match[1]));
}

export function buildNumberRange(from: number, to: number): number[] {
  const start = Math.floor(Math.min(from, to));
  const end = Math.floor(Math.max(from, to));
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

export function numericParam(input: OfficeActionInput, key: string): number | undefined {
  const value = input.params?.[key];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

export function normalizeHeadingLevel(value: unknown): number {
  return typeof value === "number" && value >= 1 && value <= 9 ? Math.floor(value) : 1;
}

export function excelChartType(value?: string): number {
  switch ((value || "column").toLowerCase()) {
    case "line": return 4;
    case "pie": return 5;
    case "bar": return 57;
    case "area": return 1;
    case "scatter": return -4169;
    default: return 51;
  }
}

export function pptChartType(value?: string): number {
  return excelChartType(value);
}

export function colorToOle(hex: string): number {
  const normalized = /^[0-9a-fA-F]{6}$/.test(hex) ? hex : "1F4E79";
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  return r + (g * 256) + (b * 65536);
}

export function defaultOutputPath(filePath: string, suffix: string): string {
  if (suffix.includes(".")) {
    return path.join(path.dirname(filePath), `${path.basename(filePath, path.extname(filePath))}-${suffix}`);
  }
  const ext = path.extname(filePath);
  return path.join(path.dirname(filePath), `${path.basename(filePath, ext)}-${suffix}${ext}`);
}

export function psLiteralText(value: string): string {
  return value.replace(/'/g, "''");
}

export function excelChangeKind(operation: string): string {
  if (operation === "insertChart") return "chart";
  if (operation === "setDataValidation") return "validation";
  if (operation === "applyConditionalFormatting") return "conditional-format";
  if (["createPivotTable", "refreshPivotTables"].includes(operation)) return "pivot-table";
  if (operation === "addSlicer") return "slicer";
  if (operation === "createPowerQuery") return "power-query";
  if (operation === "managePowerQuery") return "power-query";
  if (operation === "formatChart") return "chart-style";
  if (operation === "manageWorkbookObject") return "workbook-object";
  if (operation === "manageWorksheetObjects") return "worksheet-object";
  if (operation === "applyWorkbookTemplate") return "workbook-template";
  if (operation === "configurePrint") return "print-settings";
  if (operation === "repairFormulaReferences") return "formula-repair";
  if (operation === "convertFormulasToValues") return "formula-values";
  if (operation === "restoreFormulas") return "formula-restore";
  if (operation === "manageFormulaProtection") return "formula-protection";
  return "table-style";
}

export function excelChangeDetail(operation: string): string {
  if (operation === "insertChart") return "已创建图表";
  if (operation === "setDataValidation") return "已设置数据验证";
  if (operation === "applyConditionalFormatting") return "已应用条件格式";
  if (operation === "createPivotTable") return "已创建数据透视表";
  if (operation === "refreshPivotTables") return "已刷新数据透视表";
  if (operation === "addSlicer") return "已添加切片器";
  if (operation === "createPowerQuery") return "已创建 Power Query";
  if (operation === "managePowerQuery") return "已管理 Power Query 数据管道";
  if (operation === "formatChart") return "已格式化图表";
  if (operation === "manageWorkbookObject") return "已管理工作簿对象";
  if (operation === "manageWorksheetObjects") return "已管理工作表对象";
  if (operation === "applyWorkbookTemplate") return "已应用工作簿模板";
  if (operation === "configurePrint") return "已配置打印设置";
  if (operation === "repairFormulaReferences") return "已修复公式错误引用";
  if (operation === "convertFormulasToValues") return "已将公式转换为值";
  if (operation === "restoreFormulas") return "已恢复公式";
  if (operation === "manageFormulaProtection") return "已管理公式区域保护";
  return "已应用表格样式";
}

export function wordChangeKind(operation: string): string {
  if (operation === "insertOrUpdateToc") return "toc";
  if (operation === "insertOrReplaceImage") return "image";
  if (operation === "snapshot") return "snapshot";
  if (operation === "manageReferences") return "reference";
  if (operation === "manageRevisions") return "revision";
  if (operation === "applyTrackedChanges") return "revision";
  if (operation === "compareDocuments") return "document-compare";
  if (operation === "mailMerge") return "mail-merge";
  if (operation === "batchMailMerge") return "mail-merge";
  if (operation === "prepareMailMergeTemplate") return "mail-merge-template";
  if (operation === "populateContentControls") return "content-control";
  if (operation === "manageContentControls") return "content-control";
  return "document-style";
}

export function wordChangeDetail(operation: string): string {
  if (operation === "insertOrUpdateToc") return "已插入或更新目录";
  if (operation === "insertOrReplaceImage") return "已插入图片";
  if (operation === "snapshot") return "已导出文档预览";
  if (operation === "formatLongDocument") return "已规范长文档格式";
  if (operation === "manageReferences") return "已处理文档引用";
  if (operation === "manageRevisions") return "已处理审阅修订";
  if (operation === "applyTrackedChanges") return "已在修订模式下应用修改";
  if (operation === "compareDocuments") return "已对比文档并生成变更摘要";
  if (operation === "mailMerge") return "已生成邮件合并文档";
  if (operation === "batchMailMerge") return "已批量生成邮件合并文档";
  if (operation === "prepareMailMergeTemplate") return "已准备邮件合并模板域";
  if (operation === "populateContentControls") return "已填充内容控件";
  if (operation === "manageContentControls") return "已管理内容控件";
  return "已应用 Word 样式";
}

export function presentationChangeKind(operation: string): string {
  if (operation === "snapshot") return "snapshot";
  if (operation === "deleteSlides") return "slide-delete";
  if (operation === "insertChart") return "chart";
  if (operation === "replacePictureSlot") return "image";
  if (operation === "configureAnimations") return "animation";
  if (operation === "setSpeakerNotes") return "speaker-notes";
  if (operation === "exportHandouts") return "handout";
  return "presentation-style";
}

export function presentationChangeDetail(operation: string): string {
  if (operation === "snapshot") return "已导出幻灯片快照";
  if (operation === "deleteSlides") return "已删除幻灯片";
  if (operation === "insertChart") return "已插入图表";
  if (operation === "replacePictureSlot") return "已替换图片占位";
  if (operation === "alignShapes") return "已对齐形状";
  if (operation === "normalizeLayouts") return "已规范版式";
  if (operation === "applyMasterBranding") return "已应用母版品牌规范";
  if (operation === "layoutElements") return "已重新布局页面元素";
  if (operation === "configureAnimations") return "已配置动画";
  if (operation === "setSpeakerNotes") return "已写入演讲者备注";
  if (operation === "exportHandouts") return "已导出讲义";
  return "已应用 PPT 主题";
}
