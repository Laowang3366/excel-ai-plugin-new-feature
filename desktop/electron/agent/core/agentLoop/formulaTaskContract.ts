import type { ToolCallItem, ToolResultItem, Turn, TurnItem } from "../../shared/types";

export type FormulaReferenceMode = "none" | "partial" | "complete";

export type FormulaAcceptanceCheckType =
  | "no_excel_error"
  | "shape"
  | "unique_key"
  | "row_count"
  | "aggregate_reconciliation"
  | "sort_order"
  | "lookup_consistency"
  | "pattern_match"
  | "boundary"
  | "spot_check";

export interface FormulaRangeRef {
  raw: string;
  sheetName?: string;
  range: string;
}

export interface FormulaTaskContract {
  task: string;
  hostEnvironment?: string;
  dataSourceRanges: FormulaRangeRef[];
  referenceRange?: FormulaRangeRef;
  referenceMode: FormulaReferenceMode;
  targetRange?: FormulaRangeRef;
  targetChosenByAgent: boolean;
  sourceContent: string;
}

export interface FormulaAcceptanceCheck {
  type: FormulaAcceptanceCheckType;
  description: string;
  required: boolean;
  params: Record<string, unknown>;
}

export interface FormulaPreparation {
  status: "ready" | "needs_clarification";
  scenario: string;
  inputShape: string;
  outputShape: string;
  inputGrain: string;
  outputGrain: string;
  businessKeys: string[];
  transformChain: string[];
  constraints: string[];
  acceptanceChecks: FormulaAcceptanceCheck[];
  assumptions: string[];
  clarificationQuestion?: string;
}

export interface RangeReadEvidence {
  call: ToolCallItem;
  result: ToolResultItem;
  ref: FormulaRangeRef;
  matrix: unknown[][] | null;
  callIndex: number;
  resultIndex: number;
}

const FORMULA_MODULE_MARKERS = [
  "【功能模块：生成公式】",
  "【功能模块：公式助手】",
];

const FORMULA_WRITE_MARKERS = [
  ...FORMULA_MODULE_MARKERS,
  "写入公式",
  "公式写入",
];

export function isFormulaWorkflowTurn(turn: Turn): boolean {
  return getTurnUserContent(turn).some((content) =>
    FORMULA_WRITE_MARKERS.some((marker) => content.includes(marker)),
  );
}

export function parseFormulaTaskContract(turn: Turn): FormulaTaskContract | null {
  const content = [...getTurnUserContent(turn)]
    .reverse()
    .find((item) => FORMULA_WRITE_MARKERS.some((marker) => item.includes(marker)));
  if (!content) return null;

  const task = lineValue(content, "任务说明") || lineValue(content, "公式需求") || "生成公式";
  const dataSourceText = lineValue(content, "数据源选区");
  const referenceText = lineValue(content, "答案参考样例");
  const referenceModeText = lineValue(content, "答案参考样例类型");
  const targetText = lineValue(content, "答案填入锚点/选区");
  const targetChosenByAgent = !targetText || targetText.includes("由 Agent 选择") || targetText === "未指定";

  return {
    task,
    hostEnvironment: lineValue(content, "当前连接环境") || undefined,
    dataSourceRanges: splitRanges(dataSourceText).map(parseFormulaRangeRef).filter(isRangeRef),
    referenceRange: isMissingRange(referenceText) ? undefined : parseFormulaRangeRef(referenceText),
    referenceMode: parseReferenceMode(referenceText, referenceModeText),
    targetRange: targetChosenByAgent ? undefined : parseFormulaRangeRef(targetText),
    targetChosenByAgent,
    sourceContent: content,
  };
}

export function collectSuccessfulRangeReads(turn: Turn): RangeReadEvidence[] {
  const callsById = new Map<string, { call: ToolCallItem; index: number }>();
  turn.items.forEach((item, index) => {
    if (item.type === "tool_call" && item.toolName === "range.read") {
      callsById.set(item.id, { call: item, index });
    }
  });

  const reads: RangeReadEvidence[] = [];
  turn.items.forEach((item, resultIndex) => {
    if (item.type !== "tool_result" || item.toolName !== "range.read" || item.isError) return;
    const match = callsById.get(item.toolCallId);
    if (!match) return;
    const ref = rangeRefFromRead(match.call, item);
    if (!ref) return;
    reads.push({
      call: match.call,
      result: item,
      ref,
      matrix: matrixFromReadResult(item.result),
      callIndex: match.index,
      resultIndex,
    });
  });
  return reads;
}

export function getMissingRequiredReads(
  turn: Turn,
  contract: FormulaTaskContract,
): FormulaRangeRef[] {
  const reads = collectSuccessfulRangeReads(turn).filter(
    (read) => read.call.arguments.expand !== "spill",
  );
  const required = [...contract.dataSourceRanges];
  if (contract.referenceRange) required.push(contract.referenceRange);

  if (required.length === 0) {
    return reads.length > 0 ? [] : [{ raw: "当前数据结构", range: "当前数据结构" }];
  }
  return required.filter((expected) =>
    !reads.some((actual) => rangeRefContains(actual.ref, expected)),
  );
}

export function findReadEvidence(
  turn: Turn,
  expected: FormulaRangeRef,
): RangeReadEvidence | undefined {
  return collectSuccessfulRangeReads(turn)
    .filter((read) => rangeRefContains(read.ref, expected))
    .sort((a, b) => b.resultIndex - a.resultIndex)[0];
}

export function normalizeFormulaPreparation(args: Record<string, unknown>): FormulaPreparation | string {
  const status = args.status;
  if (status !== "ready" && status !== "needs_clarification") {
    return "status 必须是 ready 或 needs_clarification";
  }
  if (status === "needs_clarification") {
    const question = asText(args.clarificationQuestion);
    if (!question) return "needs_clarification 状态必须提供 clarificationQuestion";
    return {
      status,
      scenario: asText(args.scenario) || "待确认",
      inputShape: asText(args.inputShape) || "待确认",
      outputShape: asText(args.outputShape) || "待确认",
      inputGrain: asText(args.inputGrain) || "待确认",
      outputGrain: asText(args.outputGrain) || "待确认",
      businessKeys: asStringArray(args.businessKeys),
      transformChain: asStringArray(args.transformChain),
      constraints: asStringArray(args.constraints),
      acceptanceChecks: normalizeAcceptanceChecks(args.acceptanceChecks),
      assumptions: asStringArray(args.assumptions),
      clarificationQuestion: question,
    };
  }

  const requiredText: Array<[keyof FormulaPreparation, unknown]> = [
    ["scenario", args.scenario],
    ["inputShape", args.inputShape],
    ["outputShape", args.outputShape],
    ["inputGrain", args.inputGrain],
    ["outputGrain", args.outputGrain],
  ];
  const missing = requiredText.filter(([, value]) => !asText(value)).map(([key]) => key);
  if (missing.length > 0) return `缺少结构判断字段: ${missing.join(", ")}`;

  const transformChain = asStringArray(args.transformChain);
  if (transformChain.length === 0) return "transformChain 至少包含一个必要变换";
  const acceptanceChecks = normalizeAcceptanceChecks(args.acceptanceChecks);
  if (acceptanceChecks.length === 0) return "acceptanceChecks 至少包含一个验收检查";

  return {
    status,
    scenario: asText(args.scenario),
    inputShape: asText(args.inputShape),
    outputShape: asText(args.outputShape),
    inputGrain: asText(args.inputGrain),
    outputGrain: asText(args.outputGrain),
    businessKeys: asStringArray(args.businessKeys),
    transformChain,
    constraints: asStringArray(args.constraints),
    acceptanceChecks,
    assumptions: asStringArray(args.assumptions),
  };
}

export function getLatestFormulaPreparation(turn: Turn): {
  preparation: FormulaPreparation;
  resultIndex: number;
  toolCallId: string;
} | null {
  for (let index = turn.items.length - 1; index >= 0; index--) {
    const item = turn.items[index];
    if (item.type !== "tool_result" || item.toolName !== "formula.prepare" || item.isError) continue;
    if (!isFormulaPreparation(item.result)) continue;
    return { preparation: item.result, resultIndex: index, toolCallId: item.toolCallId };
  }
  return null;
}

export function rangeRefContains(actual: FormulaRangeRef, expected: FormulaRangeRef): boolean {
  if (expected.sheetName && actual.sheetName && normalizeSheet(actual.sheetName) !== normalizeSheet(expected.sheetName)) {
    return false;
  }
  const actualBounds = parseA1Bounds(actual.range);
  const expectedBounds = parseA1Bounds(expected.range);
  if (!actualBounds || !expectedBounds) {
    return normalizeAddress(actual.range) === normalizeAddress(expected.range);
  }
  return actualBounds.startRow <= expectedBounds.startRow
    && actualBounds.startCol <= expectedBounds.startCol
    && actualBounds.endRow >= expectedBounds.endRow
    && actualBounds.endCol >= expectedBounds.endCol;
}

export function sameFormulaAnchor(actual: FormulaRangeRef, expected: FormulaRangeRef): boolean {
  if (expected.sheetName && actual.sheetName && normalizeSheet(actual.sheetName) !== normalizeSheet(expected.sheetName)) {
    return false;
  }
  const actualBounds = parseA1Bounds(actual.range);
  const expectedBounds = parseA1Bounds(expected.range);
  if (!actualBounds || !expectedBounds) {
    return normalizeAddress(actual.range) === normalizeAddress(expected.range);
  }
  return actualBounds.startRow === expectedBounds.startRow
    && actualBounds.startCol === expectedBounds.startCol;
}

export function parseFormulaRangeRef(value: string): FormulaRangeRef | undefined {
  const raw = value.trim();
  if (!raw) return undefined;
  const separator = raw.lastIndexOf("!");
  if (separator < 0) return { raw, range: raw.replace(/\$/g, "") };
  const sheetName = raw.slice(0, separator).trim().replace(/^'|'$/g, "");
  const range = raw.slice(separator + 1).trim().replace(/\$/g, "");
  if (!range) return undefined;
  return { raw, sheetName, range };
}

function rangeRefFromRead(call: ToolCallItem, result: ToolResultItem): FormulaRangeRef | undefined {
  const sheetName = asText(call.arguments.sheetName) || undefined;
  const requestedRange = asText(call.arguments.range);
  const resultAddress = isRecord(result.result) ? asText(result.result.address) : "";
  const range = resultAddress || requestedRange;
  if (!range) return undefined;
  return { raw: sheetName ? `${sheetName}!${range}` : range, sheetName, range };
}

export function matrixFromReadResult(result: unknown): unknown[][] | null {
  const value = isRecord(result) && "values" in result ? result.values : result;
  if (!Array.isArray(value)) return [[value]];
  if (value.length === 0) return [];
  if (!Array.isArray(value[0])) return [value];
  return value.map((row) => Array.isArray(row) ? row : [row]);
}

function normalizeAcceptanceChecks(value: unknown): FormulaAcceptanceCheck[] {
  if (!Array.isArray(value)) return [];
  const supported = new Set<FormulaAcceptanceCheckType>([
    "no_excel_error",
    "shape",
    "unique_key",
    "row_count",
    "aggregate_reconciliation",
    "sort_order",
    "lookup_consistency",
    "pattern_match",
    "boundary",
    "spot_check",
  ]);
  const checks: FormulaAcceptanceCheck[] = [];
  for (const item of value) {
    if (!isRecord(item) || typeof item.type !== "string") continue;
    if (!supported.has(item.type as FormulaAcceptanceCheckType)) continue;
    checks.push({
      type: item.type as FormulaAcceptanceCheckType,
      description: asText(item.description) || item.type,
      required: item.required !== false,
      params: isRecord(item.params) ? item.params : {},
    });
  }
  return checks;
}

function parseReferenceMode(referenceText: string, modeText: string): FormulaReferenceMode {
  if (isMissingRange(referenceText)) return "none";
  return modeText.includes("完整") ? "complete" : "partial";
}

function splitRanges(value: string): string[] {
  if (isMissingRange(value)) return [];
  return value.split(/[；;]/).map((item) => item.trim()).filter(Boolean);
}

function isMissingRange(value: string): boolean {
  return !value || value === "未指定" || value.includes("自主判断");
}

function lineValue(content: string, label: string): string {
  const prefix = `${label}：`;
  const line = content.split(/\r?\n/).find((item) => item.trim().startsWith(prefix));
  return line ? line.trim().slice(prefix.length).trim() : "";
}

function getTurnUserContent(turn: Turn): string[] {
  return turn.items
    .filter((item): item is Extract<TurnItem, { type: "user_message" }> => item.type === "user_message")
    .map((item) => item.content);
}

function asText(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function asStringArray(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === "string").map((item) => item.trim()).filter(Boolean)
    : [];
}

function isRangeRef(value: FormulaRangeRef | undefined): value is FormulaRangeRef {
  return Boolean(value);
}

function isFormulaPreparation(value: unknown): value is FormulaPreparation {
  if (!isRecord(value)) return false;
  return value.status === "ready" || value.status === "needs_clarification";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function normalizeSheet(value: string): string {
  return value.trim().replace(/^'|'$/g, "").toLowerCase();
}

function normalizeAddress(value: string): string {
  return value.replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
}

function parseA1Bounds(value: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null {
  const normalized = normalizeAddress(value);
  const [startText, endText = startText] = normalized.split(":");
  const start = parseA1Cell(startText);
  const end = parseA1Cell(endText);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function parseA1Cell(value: string): { row: number; col: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(value);
  if (!match) return null;
  let col = 0;
  for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]), col };
}
