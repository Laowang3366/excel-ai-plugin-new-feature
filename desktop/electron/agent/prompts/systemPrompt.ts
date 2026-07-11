import basePrompt from "./templates/system/base.zh-CN.md?raw";
import securityPrompt from "./templates/system/security.zh-CN.md?raw";
import formulaPrompt from "./templates/scenarios/formula.zh-CN.md?raw";
import ocrInvoicePrompt from "./templates/scenarios/ocr-invoice.zh-CN.md?raw";
import officeToolsPrompt from "./templates/scenarios/office-tools.zh-CN.md?raw";
import generalOfficePrompt from "./templates/scenarios/general-office.zh-CN.md?raw";
import runtimeEnvironmentPrompt from "./templates/runtime/environment.zh-CN.md?raw";
import { composePromptSections, renderPromptTemplate } from "./promptComposer";

export { appendFolderContext } from "./sections/folderContextPrompt";
export type { FolderFileItem } from "./sections/folderContextPrompt";

type PromptAttachment = {
  fileName?: string;
  filePath?: string;
  fileType?: string;
};

export interface PromptBuildContext {
  content?: string;
  attachments?: PromptAttachment[];
  folderId?: string;
}

export interface RuntimePromptContext {
  officeConnectionStatus: string;
  dynamicArrayFunctionsEnabled: boolean;
  now?: Date;
}

interface ContextualPromptDefinition {
  key: string;
  content: string;
  shouldInclude: (context: PromptBuildContext) => boolean;
}

const baseSections = [
  { key: "base", content: basePrompt },
  { key: "security", content: securityPrompt },
];

const contextualSections: ContextualPromptDefinition[] = [
  { key: "formula", content: formulaPrompt, shouldInclude: isFormulaPromptContext },
  { key: "ocr-invoice", content: ocrInvoicePrompt, shouldInclude: shouldInjectOcrRules },
  {
    key: "office-tools",
    content: officeToolsPrompt,
    shouldInclude: (context) => !isFormulaPromptContext(context) && shouldInjectOfficeTools(context),
  },
  {
    key: "general-office",
    content: generalOfficePrompt,
    shouldInclude: (context) => !isFormulaPromptContext(context) && shouldInjectGeneralScenarios(context),
  },
];

export function buildSystemPrompt(): string {
  return composePromptSections(baseSections);
}

export function buildContextualPromptSections(context: PromptBuildContext = {}): string {
  return composePromptSections(
    contextualSections
      .filter((section) => section.shouldInclude(context))
      .map(({ key, content }) => ({ key, content })),
  );
}

export function buildRuntimePromptSection(context: RuntimePromptContext): string {
  const now = context.now ?? new Date();
  const dateFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "long",
  });
  const timeFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  const dynamicArraySupport = context.dynamicArrayFunctionsEnabled
    ? "已开启。生成 Excel/WPS 公式时默认允许 FILTER、UNIQUE、SORT、SEQUENCE、LET、XLOOKUP 等动态数组函数；不要反复质疑当前环境是否适配动态数组函数，除非工具回读明确返回 #NAME? 或用户关闭此设置。"
    : "已关闭。生成 Excel/WPS 公式时不要依赖动态数组 spill，优先使用逐格独立公式、传统函数或辅助区域。";
  return renderPromptTemplate(runtimeEnvironmentPrompt, {
    OFFICE_CONNECTION_CONTEXT: `- Office 应用连接状态：${context.officeConnectionStatus}`,
    DYNAMIC_ARRAY_SUPPORT: dynamicArraySupport,
    CURRENT_DATE: dateFormatter.format(now),
    CURRENT_TIME: timeFormatter.format(now),
  });
}

export function isFormulaPromptContext(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  return hasAny(content, [
    "【功能模块：公式助手】",
    "【功能模块：生成公式】",
    "range.write",
    'expand:"spill"',
    "expand:'spill'",
    "动态数组",
    "数组公式",
    "公式写入",
    "写入公式",
    "生成公式",
    "excel 公式",
    "wps 公式",
    "公式函数",
    "函数公式",
    "公式场景",
    "公式助手",
    "spill",
    "#spill",
  ]);
}

function shouldInjectOcrRules(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  if (
    hasAny(content, [
      "【功能模块：发票识别】",
      "发票识别",
      "ocr",
      "识别字段",
      "字段识别",
      "图片识别",
      "图片解析",
      "票据识别",
      "ocr.parsedocument",
    ])
  ) {
    return true;
  }
  return context.attachments?.some((attachment) => isImageOrPdfAttachment(attachment)) ?? false;
}

function shouldInjectOfficeTools(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  if (context.attachments?.some((attachment) => isOfficeAttachment(attachment)) ?? false) {
    return true;
  }
  return hasAny(content, [
    ".xlsx",
    ".xls",
    ".csv",
    ".docx",
    ".doc",
    ".pptx",
    ".ppt",
    "open xml",
    "office.action",
    "excel",
    "wps",
    "word",
    "ppt",
    "powerpoint",
    "表格",
    "单元格",
    "选区",
    "工作簿",
    "工作表",
    "当前表格",
    "当前工作簿",
    "当前工作表",
    "公式",
    "当前文件",
    "当前文档",
    "演示文稿",
    "幻灯片",
    "美化",
    "版面",
    "视觉设计",
    "样式",
    "文件级创建",
    "文件级编辑",
    "打开文件",
    "保存工作簿",
    "保存文档",
    "保存演示文稿",
    "校验表格",
    "验证表格",
  ]);
}

function shouldInjectGeneralScenarios(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  return hasAny(content, [
    "数据清洗",
    "清洗",
    "图表",
    "报告",
    "汇总",
    "统计",
    "趋势",
    "批量",
    "条件格式",
    "数据验证",
    "建模",
    "预测",
    "vba",
    "脚本",
    "宏",
  ]);
}

function normalizeContent(content: string | undefined): string {
  return (content ?? "").toLowerCase();
}

function hasAny(content: string, needles: string[]): boolean {
  return needles.some((needle) => content.includes(needle.toLowerCase()));
}

function isImageOrPdfAttachment(attachment: PromptAttachment): boolean {
  const name = `${attachment.fileName ?? ""} ${attachment.filePath ?? ""}`.toLowerCase();
  return attachment.fileType === "image" || /\.(png|jpe?g|webp|bmp|gif|tiff?|pdf)$/i.test(name);
}

function isOfficeAttachment(attachment: PromptAttachment): boolean {
  const name = `${attachment.fileName ?? ""} ${attachment.filePath ?? ""}`.toLowerCase();
  return /\.(xlsx|xlsm?|xlsb|csv|docx?|pptx?)$/i.test(name);
}
