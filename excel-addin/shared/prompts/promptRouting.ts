/**
 * Excel-only prompt routing adapted from desktop promptRouting.
 * Word/PPT-only intents stay excluded; OCR is add-in vision path (no parseDocument).
 */

export type ExcelPromptScenario = "formula" | "office-tools" | "general-office" | "macro" | "ocr-invoice";

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
const OCR_INTENT_PATTERN =
  /(?:发票|票据|图片|字段).{0,6}(?:识别|解析)|(?:识别|解析).{0,6}(?:发票|票据|图片|字段)|\bocr\b|OCR识别|功能模块：\s*OCR/iu;
const FORMULA_INTENT_PATTERN =
  /(?:动态数组|数组公式|公式(?:写入|生成|函数|场景|助手)|(?:写入|生成|编写|创建).{0,8}(?:excel\s*|wps\s*)?公式|range\.write|expand\s*:\s*["']spill["']|#?spill)/iu;
const EXCEL_INTENT_PATTERN =
  /(?:excel|wps|表格|单元格|选区|工作簿|工作表|office\.action|\.xlsx?|\.csv)/iu;
const GENERAL_OFFICE_INTENT_PATTERN =
  /(?:数据)?清洗|图表|报告|汇总|统计|趋势|批量|条件格式|数据验证|建模|预测/iu;
const MACRO_INTENT_PATTERN =
  /(?:创建|编写|写入|安装|修改|修复|绑定|运行|执行).{0,10}(?:vba|宏|jsa|脚本按钮|控制按钮)|(?:vba|宏|jsa|脚本按钮|控制按钮).{0,10}(?:创建|编写|写入|安装|修改|修复|绑定|运行|执行)|点击.{0,8}按钮|按钮.{0,8}(?:点击|绑定|切换)|application\.caller|onaction|macro\.write/iu;
const EXCEL_ATTACHMENT_PATTERN = /\.(?:xlsx|xlsm?|xlsb|csv)$/iu;
const POWER_QUERY_EXPLICIT_PATTERN = /(?:power\s*query|powerquery|查询编辑器|m\s*公式)/iu;
const REFRESHABLE_ETL_PATTERN =
  /(?:(?:外部|多来源|多数据源|external|multi[ -]?source|multiple\s+sources).{0,16}(?:可刷新|自动刷新|定时刷新|刷新|refresh(?:able)?|etl|提取|转换|加载)|(?:可刷新|自动刷新|定时刷新|刷新|refresh(?:able)?|etl).{0,16}(?:外部|多来源|多数据源|external|multi[ -]?source|multiple\s+sources))/iu;
const INTERACTIVE_PIVOT_PATTERN =
  /(?:数据透视表|透视表|pivot\s*table|pivottable|切片器|slicer|交互式?透视|透视分析)/iu;

/** Word/PPT-only intents stay excluded; ocr-invoice is add-in-adapted (no ocr.parseDocument). */
export const EXCLUDED_SCENARIOS = [] as const;

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

export function resolveExcelPromptScenarios(
  context: PromptRoutingContext,
): Set<ExcelPromptScenario> {
  const content = context.content ?? "";
  const moduleName = FEATURE_MODULE_PATTERN.exec(content)?.[1]?.trim() ?? "";
  const scenarios = new Set<ExcelPromptScenario>();
  const advancedOfficeIntents = resolveOfficeAdvancedIntents(context);

  const isFormula = moduleName.includes("公式") || FORMULA_INTENT_PATTERN.test(content);
  if (isFormula) {
    scenarios.add("formula");
  }

  const moduleIsOcr =
    moduleName.includes("OCR") ||
    moduleName.includes("ocr") ||
    moduleName.includes("发票") ||
    moduleName.includes("识别");
  if (
    moduleIsOcr ||
    OCR_INTENT_PATTERN.test(content) ||
    context.attachments?.some((a) =>
      /\.(?:png|jpe?g|webp|bmp|gif|pdf)$/iu.test(a.fileName || a.filePath || ""),
    )
  ) {
    scenarios.add("ocr-invoice");
  }

  const isMacro = MACRO_INTENT_PATTERN.test(content);
  if (isMacro) {
    scenarios.add("macro");
  }

  if (!isFormula) {
    if (
      !isMacro &&
      (moduleName.includes("代码") ||
        moduleName.includes("报告") ||
        moduleName.includes("清洗") ||
        moduleName.includes("图表") ||
        advancedOfficeIntents.size > 0 ||
        EXCEL_INTENT_PATTERN.test(content) ||
        context.attachments?.some(isExcelAttachment))
    ) {
      scenarios.add("office-tools");
    }

    if (moduleName.includes("报告") ||
        moduleName.includes("清洗") ||
        moduleName.includes("图表") || GENERAL_OFFICE_INTENT_PATTERN.test(content)) {
      scenarios.add("general-office");
    }
  }

  return scenarios;
}

function getAttachmentName(attachment: PromptAttachment): string {
  return attachment.filePath || attachment.fileName || "";
}

function isExcelAttachment(attachment: PromptAttachment): boolean {
  return EXCEL_ATTACHMENT_PATTERN.test(getAttachmentName(attachment));
}
