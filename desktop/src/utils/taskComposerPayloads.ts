export type HostEnvironment = "wps" | "microsoft_excel" | "unknown";
export type ReferenceSampleMode = "partial" | "complete";
export type PreferredLanguage = "auto" | "js" | "vba" | "python";
export type ReportOutputFormat = "excel" | "word" | "ppt";

export interface ExcelStatusLike {
  connected: boolean;
  host: string;
  version?: string;
  workbookName?: string;
}

export interface RangeDraftInput {
  dataSourceRanges: string[];
  dataSourceInput: string;
}

export interface FormulaPayloadInput extends RangeDraftInput {
  referenceSampleRange: string;
  referenceSampleMode: ReferenceSampleMode;
  outputRange: string;
  hostEnvironment: HostEnvironment;
  task: string;
}

export interface CodePayloadInput extends RangeDraftInput {
  referenceSampleRange: string;
  referenceSampleMode: ReferenceSampleMode;
  outputRange: string;
  hostEnvironment: HostEnvironment;
  preferredLanguage: PreferredLanguage;
  task: string;
}

export interface ReportPayloadInput {
  range: string;
  task: string;
  outputFormat: ReportOutputFormat;
  storagePath: string;
}

export const REFERENCE_SAMPLE_MODE_LABELS: Record<ReferenceSampleMode, string> = {
  partial: "部分样例",
  complete: "完整样例",
};

export const REPORT_OUTPUT_FORMAT_LABELS: Record<ReportOutputFormat, string> = {
  excel: "Excel",
  word: "Word 文档",
  ppt: "PPT",
};

export function mergeRangeInput({ dataSourceRanges, dataSourceInput }: RangeDraftInput): string[] {
  const trimmedInput = dataSourceInput.trim();
  if (trimmedInput && !dataSourceRanges.includes(trimmedInput)) {
    return [...dataSourceRanges, trimmedInput];
  }
  return dataSourceRanges;
}

export function normalizeHostEnvironment(status?: ExcelStatusLike | null): HostEnvironment {
  const host = status?.host?.toLowerCase() ?? "";
  if (!status?.connected && !host) return "unknown";
  if (host.includes("wps") || host.includes("ket")) return "wps";
  if (host.includes("excel") || host.includes("office")) return "microsoft_excel";
  return "unknown";
}

export function getHostEnvironmentLabel(environment: HostEnvironment): string {
  if (environment === "wps") return "WPS";
  if (environment === "microsoft_excel") return "Microsoft Excel";
  return "未检测到连接环境";
}

export function buildFormulaTaskPayload(input: FormulaPayloadInput): string {
  const lines: string[] = ["【功能模块：生成公式】"];
  const taskText = input.task.trim() || "生成公式";
  const dataSourceValues = mergeRangeInput(input);

  lines.push(`任务说明：${taskText}`);
  lines.push(`当前连接环境：${getHostEnvironmentLabel(input.hostEnvironment)}`);
  lines.push("交付要求：必须使用 Excel/WPS 函数公式完成，不要改写为 VBA、JS、Python 或手工值。");
  if (dataSourceValues.length > 0) {
    lines.push(`数据源选区：${dataSourceValues.join("；")}`);
  }
  if (input.referenceSampleRange.trim()) {
    lines.push(`答案参考样例：${input.referenceSampleRange.trim()}`);
    lines.push(`答案参考样例类型：${REFERENCE_SAMPLE_MODE_LABELS[input.referenceSampleMode]}`);
  }
  if (input.outputRange.trim()) {
    lines.push(`答案填入锚点/选区：${input.outputRange.trim()}`);
  } else {
    lines.push("答案填入锚点/选区：由 Agent 选择空白区域");
  }
  return lines.join("\n");
}

export function buildCodeTaskPayload(input: CodePayloadInput): string {
  const taskText = input.task.trim() || "用适合当前环境的代码方案解决工作簿需求。";
  const dataSourceValues = mergeRangeInput(input);
  const languageLabel: Record<PreferredLanguage, string> = {
    auto: "自动",
    js: "JS",
    vba: "VBA",
    python: "Python",
  };

  const lines = [
    "【功能模块：代码生成】",
    `代码需求：${taskText}`,
    `运行环境：${getHostEnvironmentLabel(input.hostEnvironment)}`,
    `首选语言：${languageLabel[input.preferredLanguage]}`,
  ];
  if (dataSourceValues.length > 0) {
    lines.push(`数据源选区：${dataSourceValues.join("；")}`);
  } else {
    lines.push("数据源选区：未指定，请读取工作簿快照后自主判断。");
  }
  if (input.referenceSampleRange.trim()) {
    lines.push(`答案参考样例：${input.referenceSampleRange.trim()}`);
    lines.push(`答案参考样例类型：${REFERENCE_SAMPLE_MODE_LABELS[input.referenceSampleMode]}`);
  } else {
    lines.push("答案参考样例：未指定。");
  }
  if (input.outputRange.trim()) {
    lines.push(`输出/操作锚点：${input.outputRange.trim()}`);
  } else {
    lines.push("输出/操作锚点：未指定。");
  }

  return lines.join("\n");
}

export function buildReportTaskPayload(input: ReportPayloadInput): string {
  const formatLabel = REPORT_OUTPUT_FORMAT_LABELS[input.outputFormat];
  const lines = [
    "【功能模块：报告生成】",
    `报告类型：${formatLabel}`,
    `数据源选区：${input.range.trim() || "未指定，请读取工作簿快照后自主判断。"}`,
    `需求说明：${input.task.trim() || "基于选取数据生成结构化分析报告。"}`,
  ];

  if (input.outputFormat === "excel") {
    lines.push("交付方式：在当前连接的 Excel/WPS 环境中新增或更新报告工作表。");
  } else {
    const fileTarget = input.outputFormat === "word" ? "Word 文档" : `${formatLabel} 文件`;
    lines.push(`存储路径：${input.storagePath.trim() || "桌面"}`);
    lines.push(`交付方式：在上述路径创建 ${fileTarget}，写入报告内容后用系统默认应用打开。`);
  }

  return lines.join("\n");
}
