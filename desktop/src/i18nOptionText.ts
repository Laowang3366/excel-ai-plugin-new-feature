import type { IntentKind } from "./components/Sidebar";
import type { PermissionMode } from "./store/settingsStore";

type IntentKey = NonNullable<IntentKind>;
type SimpleTaskIntent = "clean" | "report" | "chart";

export const ZH_INTENT_LABELS = {
  formula: "公式助手",
  code: "代码生成",
  ocr: "OCR 识别",
  clean: "数据清洗",
  report: "报告生成",
  chart: "图表制作",
  office: "Office 自动化",
} satisfies Record<IntentKey, string>;

export const EN_INTENT_LABELS = {
  formula: "Formula assistant",
  code: "Code generation",
  ocr: "OCR",
  clean: "Data cleaning",
  report: "Report generation",
  chart: "Chart creation",
  office: "Office automation",
} satisfies Record<IntentKey, string>;

export const ZH_PERMISSION_LABELS = {
  normal: "逐次确认",
  auto_approve_safe: "自动批准",
  confirm_all: "完整权限（自动执行）",
} satisfies Record<PermissionMode, string>;

export const EN_PERMISSION_LABELS = {
  normal: "Confirm each time",
  auto_approve_safe: "Auto approve",
  confirm_all: "Full access (auto execute)",
} satisfies Record<PermissionMode, string>;

export const ZH_SIMPLE_PLACEHOLDERS = {
  clean: "如：去除重复行、统一日期格式、补全缺失值...",
  report: "如：按季度汇总销售数据，生成趋势分析...",
  chart: "如：根据销量和利润数据生成柱状图...",
} satisfies Record<SimpleTaskIntent, string>;

export const EN_SIMPLE_PLACEHOLDERS = {
  clean: "e.g. remove duplicates, normalize dates, fill missing values...",
  report: "e.g. summarize quarterly sales and generate trend analysis...",
  chart: "e.g. create a bar chart from sales and profit data...",
} satisfies Record<SimpleTaskIntent, string>;

export const ZH_SIMPLE_PREFIXES = {
  clean: "请帮我清洗数据：",
  report: "请帮我生成分析报告：",
  chart: "请帮我构建图表：",
} satisfies Record<SimpleTaskIntent, string>;

export const EN_SIMPLE_PREFIXES = {
  clean: "Please clean this data:",
  report: "Please generate an analysis report:",
  chart: "Please create a chart:",
} satisfies Record<SimpleTaskIntent, string>;

export const ZH_TIME_TEXT = {
  seconds: (value: number) => `${value}秒`,
  minuteSecond: (minutes: number, seconds: number) =>
    seconds > 0 ? `${minutes}分${seconds}秒` : `${minutes}分钟`,
  hourParts: (hours: number, minutes: number, seconds: number) => {
    const parts = [`${hours}小时`];
    if (minutes > 0) parts.push(`${minutes}分`);
    if (seconds > 0) parts.push(`${seconds}秒`);
    return parts.join("");
  },
};

export const EN_TIME_TEXT = {
  seconds: (value: number) => `${value}s`,
  minuteSecond: (minutes: number, seconds: number) =>
    seconds > 0 ? `${minutes}m ${seconds}s` : `${minutes}m`,
  hourParts: (hours: number, minutes: number, seconds: number) => {
    const parts = [`${hours}h`];
    if (minutes > 0) parts.push(`${minutes}m`);
    if (seconds > 0) parts.push(`${seconds}s`);
    return parts.join(" ");
  },
};
