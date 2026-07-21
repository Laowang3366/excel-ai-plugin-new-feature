/** Desktop-aligned task payload builders for the Excel add-in (no IPC/paths). */

import { ocrResultProtocolInstruction } from "./ocrResultProtocol";

export type HostEnvironment = "wps" | "microsoft_excel" | "unknown";
export type ReferenceSampleMode = "partial" | "complete";
export type ReportOutputFormat = "excel" | "word" | "ppt";
export type CleanOpMode = "drop_empty" | "dedupe" | "normalize";

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

export interface CleanPayloadInput {
  range: string;
  task: string;
  modes: CleanOpMode[];
}

export interface ChartPayloadInput {
  range: string;
  task: string;
  chartType: string;
  title: string;
  showLegend: boolean;
  positionNote: string;
}

export interface ReportPayloadInput {
  range: string;
  task: string;
  outputFormat: ReportOutputFormat;
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

export const CLEAN_MODE_LABELS: Record<CleanOpMode, string> = {
  drop_empty: "去空行/空值",
  dedupe: "去重",
  normalize: "规范化/格式整理",
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
  if (host.includes("wps") || host.includes("ket") || host.includes("jsa")) return "wps";
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
  lines.push(
    "交付要求：必须使用 Excel/WPS 函数公式完成，不要改写为 VBA、JS、Python 或手工值。",
  );
  if (dataSourceValues.length > 0) {
    lines.push(`数据源选区：${dataSourceValues.join("；")}`);
  }
  if (input.referenceSampleRange.trim()) {
    lines.push(`答案参考样例：${input.referenceSampleRange.trim()}`);
    lines.push(
      `答案参考样例类型：${REFERENCE_SAMPLE_MODE_LABELS[input.referenceSampleMode]}`,
    );
  }
  if (input.outputRange.trim()) {
    lines.push(`答案填入锚点/选区：${input.outputRange.trim()}`);
  } else {
    lines.push("答案填入锚点/选区：由 Agent 选择空白区域");
  }
  return lines.join("\n");
}

export function buildCleanTaskPayload(input: CleanPayloadInput): string {
  const modeText =
    input.modes.length > 0
      ? input.modes.map((m) => CLEAN_MODE_LABELS[m]).join("、")
      : "按需求自行选择去空/去重/规范化等安全操作";
  return [
    "【功能模块：数据清洗】",
    `数据源选区：${input.range.trim() || "未指定，请读取当前选区或工作簿后判断。"}`,
    `操作模式：${modeText}`,
    `清洗要求：${input.task.trim() || "清洗数据并写回当前工作簿。"}`,
    "边界：仅使用当前工作簿 range.read/write/clear 等工具；不要使用 Power Query / 外部 ETL / 磁盘路径。批量写入须走审批。",
  ].join("\n");
}

export function buildChartTaskPayload(input: ChartPayloadInput): string {
  const lines = [
    "【功能模块：图表制作】",
    `数据源选区：${input.range.trim() || "未指定，请读取当前选区后判断。"}`,
    `图表类型：${input.chartType.trim() || "自动选择合适类型"}`,
  ];
  if (input.title.trim()) lines.push(`标题：${input.title.trim()}`);
  lines.push(`图例：${input.showLegend ? "显示" : "隐藏"}`);
  if (input.positionNote.trim()) lines.push(`位置/布局：${input.positionNote.trim()}`);
  lines.push(`需求说明：${input.task.trim() || "基于数据源创建图表。"}`);
  lines.push(
    "边界：优先 chart.* 工具；若宿主返回 unsupported（例如部分 WPS JSA 无图表合同），明确告知用户，禁止猜测未验证 API。",
  );
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
    lines.push("边界：加载项不支持 Word/PPT 文件输出与磁盘路径；仅当前工作簿。");
  } else {
    lines.push(
      "交付方式：不适用于加载项（Word/PPT/磁盘路径 unsupported）。请改用 Excel 报告工作表，或改用桌面端。",
    );
  }
  return lines.join("\n");
}

export function buildOcrTaskPayload(input: {
  mode: "image" | "invoice";
  fileNames: string[];
  task: string;
  outputRange: string;
}): string {
  const lines = [
    "【功能模块：OCR识别】",
    `识别模式：${input.mode === "invoice" ? "发票" : "通用图片"}`,
    `附件：${input.fileNames.length > 0 ? input.fileNames.join("；") : "无"}`,
    `需求说明：${input.task.trim() || "识别图片中的文字/字段并整理结果。"}`,
  ];
  if (input.outputRange.trim()) {
    lines.push(`写入锚点（可选）：${input.outputRange.trim()}`);
  }
  lines.push(
    "边界：加载项无 ocr.parseDocument / 文件路径 IPC；请直接阅读本轮多模态图片内容。PDF 若无法可靠解析须返回 unsupported，禁止假成功。写入工作表由加载项 UI 经审批调用 range.write，本轮请勿自行 range.write。",
  );
  lines.push(ocrResultProtocolInstruction());
  return lines.join("\n");
}
