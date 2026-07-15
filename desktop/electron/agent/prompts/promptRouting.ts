export type PromptScenario = "formula" | "ocr-invoice" | "office-tools" | "general-office" | "macro";

export interface PromptAttachment {
  fileName?: string;
  filePath?: string;
  fileType?: string;
}

export interface PromptRoutingContext {
  content?: string;
  attachments?: PromptAttachment[];
}

export type OfficeAdvancedIntent = "refreshable-etl" | "interactive-pivot";

const FEATURE_MODULE_PATTERN = /【功能模块：\s*([^】]+?)\s*】/iu;
const FORMULA_INTENT_PATTERN =
  /(?:动态数组|数组公式|公式(?:写入|生成|函数|场景|助手)|(?:写入|生成|编写|创建).{0,8}(?:excel\s*|wps\s*)?公式|range\.write|expand\s*:\s*["']spill["']|#?spill)/iu;
const OCR_INTENT_PATTERN =
  /(?:发票|票据|图片|字段).{0,6}(?:识别|解析)|(?:识别|解析).{0,6}(?:发票|票据|图片|字段)|ocr(?:\.parsedocument)?/iu;
const OFFICE_INTENT_PATTERN =
  /(?:excel|wps|word|ppt|powerpoint|表格|单元格|选区|工作簿|工作表|文档|演示文稿|幻灯片|open\s+xml|office\.action|\.xlsx?|\.csv|\.docx?|\.pptx?)/iu;
const GENERAL_OFFICE_INTENT_PATTERN =
  /(?:数据)?清洗|图表|报告|汇总|统计|趋势|批量|条件格式|数据验证|建模|预测/iu;
const MACRO_INTENT_PATTERN =
  /(?:创建|编写|写入|安装|修改|修复|绑定|运行|执行).{0,10}(?:vba|宏|jsa|脚本按钮|控制按钮)|(?:vba|宏|jsa|脚本按钮|控制按钮).{0,10}(?:创建|编写|写入|安装|修改|修复|绑定|运行|执行)|点击.{0,8}按钮|按钮.{0,8}(?:点击|绑定|切换)|application\.caller|onaction|macro\.write/iu;
const IMAGE_OR_PDF_ATTACHMENT_PATTERN = /\.(?:png|jpe?g|webp|bmp|gif|tiff?|pdf)$/iu;
const OFFICE_ATTACHMENT_PATTERN = /\.(?:xlsx|xlsm?|xlsb|csv|docx?|pptx?)$/iu;
const POWER_QUERY_EXPLICIT_PATTERN = /(?:power\s*query|powerquery|查询编辑器|m\s*公式)/iu;
const REFRESHABLE_ETL_PATTERN =
  /(?:(?:外部|多来源|多数据源|external|multi[ -]?source|multiple\s+sources).{0,16}(?:可刷新|自动刷新|定时刷新|刷新|refresh(?:able)?|etl|提取|转换|加载)|(?:可刷新|自动刷新|定时刷新|刷新|refresh(?:able)?|etl).{0,16}(?:外部|多来源|多数据源|external|multi[ -]?source|multiple\s+sources))/iu;
const INTERACTIVE_PIVOT_PATTERN =
  /(?:数据透视表|透视表|pivot\s*table|pivottable|切片器|slicer|交互式?透视|透视分析)/iu;

export function resolveOfficeAdvancedIntents(
  context: PromptRoutingContext = {},
): Set<OfficeAdvancedIntent> {
  const content = context.content ?? "";
  const intents = new Set<OfficeAdvancedIntent>();
  if (POWER_QUERY_EXPLICIT_PATTERN.test(content) || REFRESHABLE_ETL_PATTERN.test(content)) {
    intents.add("refreshable-etl");
  }
  if (INTERACTIVE_PIVOT_PATTERN.test(content)) {
    intents.add("interactive-pivot");
  }
  return intents;
}

export function resolvePromptScenarios(context: PromptRoutingContext): Set<PromptScenario> {
  const content = context.content ?? "";
  const moduleName = FEATURE_MODULE_PATTERN.exec(content)?.[1]?.trim() ?? "";
  const scenarios = new Set<PromptScenario>();
  const advancedOfficeIntents = resolveOfficeAdvancedIntents(context);

  const isFormula = moduleName.includes("公式") || FORMULA_INTENT_PATTERN.test(content);
  if (isFormula) {
    scenarios.add("formula");
  }

  if (
    moduleName.includes("发票") ||
    OCR_INTENT_PATTERN.test(content) ||
    context.attachments?.some(isImageOrPdfAttachment)
  ) {
    scenarios.add("ocr-invoice");
  }

  const isMacro = MACRO_INTENT_PATTERN.test(content);
  if (isMacro) {
    scenarios.add("macro");
  }

  if (!isFormula) {
    if (
      !isMacro && (moduleName.includes("代码") ||
      moduleName.includes("报告") ||
      advancedOfficeIntents.size > 0 ||
      OFFICE_INTENT_PATTERN.test(content) ||
      context.attachments?.some(isOfficeAttachment))
    ) {
      scenarios.add("office-tools");
    }

    if (moduleName.includes("报告") || GENERAL_OFFICE_INTENT_PATTERN.test(content)) {
      scenarios.add("general-office");
    }
  }

  return scenarios;
}

function getAttachmentName(attachment: PromptAttachment): string {
  return attachment.filePath || attachment.fileName || "";
}

function isImageOrPdfAttachment(attachment: PromptAttachment): boolean {
  return attachment.fileType === "image" || IMAGE_OR_PDF_ATTACHMENT_PATTERN.test(getAttachmentName(attachment));
}

function isOfficeAttachment(attachment: PromptAttachment): boolean {
  return OFFICE_ATTACHMENT_PATTERN.test(getAttachmentName(attachment));
}
