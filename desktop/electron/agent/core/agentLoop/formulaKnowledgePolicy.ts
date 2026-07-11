import type { ToolCallItem, ToolExecutor, ToolResultItem, Turn } from "../../shared/types";
import {
  collectSuccessfulRangeReads,
  getLatestFormulaPreparation,
  getMissingRequiredReads,
  isFormulaWorkflowTurn,
  parseFormulaRangeRef,
  parseFormulaTaskContract,
  sameFormulaAnchor,
  type FormulaPreparation,
} from "./formulaTaskContract";
import {
  getLatestFormulaVerification,
  getLatestSuccessfulFormulaWrite,
  verifyLatestFormulaWrite,
} from "./formulaVerification";

const KNOWLEDGE_SEARCH_TOOL = "knowledge.search";
const FORMULA_PREPARE_TOOL = "formula.prepare";
const FORMULA_VERIFY_TOOL = "formula.verify";
const RANGE_READ_TOOL = "range.read";
const RANGE_WRITE_TOOL = "range.write";
const MAX_QUERY_CONTEXT_CHARS = 1_200;

export interface FormulaWorkflowDiagnostic {
  stage: "input_read" | "prepare" | "methodology" | "pre_write" | "post_write" | "completion";
  code: string;
  message: string;
  evidence?: unknown;
  nextAction: {
    tool?: string;
    arguments?: Record<string, unknown>;
    instruction: string;
  };
  retryable: boolean;
}

export function shouldLoadFormulaMethodology(
  turn: Turn,
  executors: Map<string, ToolExecutor>,
): boolean {
  const preparation = getLatestFormulaPreparation(turn);
  return isFormulaWorkflowTurn(turn)
    && preparation?.preparation.status === "ready"
    && executors.has(KNOWLEDGE_SEARCH_TOOL)
    && !findKnowledgeResult(turn, "formula_methodology", preparation.resultIndex, true);
}

export function shouldSearchFormulaScene(
  turn: Turn,
  executors: Map<string, ToolExecutor>,
): boolean {
  const preparation = getLatestFormulaPreparation(turn);
  return isFormulaWorkflowTurn(turn)
    && preparation?.preparation.status === "ready"
    && executors.has(KNOWLEDGE_SEARCH_TOOL)
    && !findKnowledgeResult(turn, "formula_scene", preparation.resultIndex, false);
}

export function buildFormulaKnowledgeQuery(turn: Turn): string {
  const contract = parseFormulaTaskContract(turn);
  const preparation = getLatestFormulaPreparation(turn)?.preparation;
  const readContext = buildReadShapeContext(turn);
  return [
    "Excel/WPS 公式解题方法论",
    contract?.task ? `当前任务：${contract.task}` : "",
    preparation ? formatPreparationForQuery(preparation) : "",
    readContext ? `已读取结构：${readContext}` : "",
    "返回完成当前变换链所需的核心方法、约束和验收要求。",
  ].filter(Boolean).join("\n").slice(0, MAX_QUERY_CONTEXT_CHARS * 2);
}

export function buildFormulaSceneQuery(turn: Turn): string {
  const contract = parseFormulaTaskContract(turn);
  const preparation = getLatestFormulaPreparation(turn)?.preparation;
  return [
    contract?.task ? `任务：${contract.task}` : "公式任务",
    preparation ? formatPreparationForQuery(preparation) : "",
    "检索具有相同场景、输入输出形状、业务键或变换链的局部经验；没有可靠匹配时返回无匹配，不使用近似案例替代当前任务判断。",
  ].filter(Boolean).join("\n").slice(0, MAX_QUERY_CONTEXT_CHARS * 2);
}

export function guardFormulaWorkflowExecutors(
  executors: Map<string, ToolExecutor>,
  turn: Turn,
): Map<string, ToolExecutor> {
  if (!isFormulaWorkflowTurn(turn)) return executors;
  const guarded = new Map(executors);
  guardFormulaPreparation(guarded, turn);
  guardFormulaRangeWrite(guarded, turn);
  guardFormulaVerification(guarded, turn);
  return guarded;
}

export function getFormulaCompletionDiagnostic(turn: Turn): FormulaWorkflowDiagnostic | null {
  if (!isFormulaWorkflowTurn(turn)) return null;
  const preparation = getLatestFormulaPreparation(turn);
  if (preparation?.preparation.status === "needs_clarification") return null;

  const latestWrite = getLatestSuccessfulFormulaWrite(turn);
  const verification = latestWrite ? getLatestFormulaVerification(turn, latestWrite.call.id) : null;
  if (verification?.status === "passed" || verification?.status === "passed_with_assumptions") {
    return null;
  }
  if (verification?.status === "failed") {
    return diagnostic(
      "completion",
      "FORMULA_ACCEPTANCE_FAILED",
      verification.summary,
      { checks: verification.checks.filter((check) => check.required && check.status !== "passed"), nextActions: verification.nextActions },
      "根据失败证据修正公式，重新写入后再次验收。",
      FORMULA_VERIFY_TOOL,
    );
  }

  const contract = parseFormulaTaskContract(turn);
  if (!contract) {
    return diagnostic("completion", "FORMULA_CONTRACT_MISSING", "无法识别公式任务契约。", undefined, "重新提交公式任务。", undefined, false);
  }
  const missingReads = getMissingRequiredReads(turn, contract);
  if (missingReads.length > 0) {
    return diagnostic(
      "input_read",
      "REQUIRED_RANGE_NOT_READ",
      "数据源或指定参考样例尚未读取完整。",
      { missingRanges: missingReads.map((item) => item.raw) },
      "调用 range.read 读取所有缺失区域。",
      RANGE_READ_TOOL,
    );
  }
  if (!preparation) {
    return diagnostic(
      "prepare",
      "FORMULA_PREPARATION_MISSING",
      "尚未提交场景、形状、变换链和验收契约。",
      { referenceMode: contract.referenceMode },
      "根据已读取数据调用 formula.prepare；存在实质歧义时使用 needs_clarification。",
      FORMULA_PREPARE_TOOL,
    );
  }
  if (!hasSuccessfulMethodologyAfterPreparation(turn, preparation.resultIndex)) {
    return diagnostic(
      "methodology",
      "METHODOLOGY_NOT_LOADED",
      "内置公式方法论尚未成功读取。",
      undefined,
      "重新调用 knowledge.search，scope 使用 formula_methodology。",
      KNOWLEDGE_SEARCH_TOOL,
    );
  }
  if (!getLatestSuccessfulFormulaWrite(turn)) {
    return diagnostic(
      "pre_write",
      "FORMULA_NOT_WRITTEN",
      "公式尚未写入目标锚点。",
      undefined,
      "根据方法论和验收契约生成公式，然后调用 range.write。",
      RANGE_WRITE_TOOL,
    );
  }
  return diagnostic(
    "post_write",
    "FORMULA_NOT_VERIFIED",
    "公式已写入，但实际结果尚未完成回读验收。",
    undefined,
    "从写入锚点使用 range.read(expand=spill) 回读，并执行 formula.verify。",
    RANGE_READ_TOOL,
  );
}

export function formatFormulaDiagnostic(value: FormulaWorkflowDiagnostic): string {
  return `公式流程校验未通过：\n${JSON.stringify(value, null, 2)}`;
}

function guardFormulaPreparation(executors: Map<string, ToolExecutor>, turn: Turn): void {
  const original = executors.get(FORMULA_PREPARE_TOOL);
  if (!original) return;
  executors.set(FORMULA_PREPARE_TOOL, {
    name: FORMULA_PREPARE_TOOL,
    async execute(args, context) {
      const contract = parseFormulaTaskContract(turn);
      if (!contract) {
        return blocked(diagnostic("prepare", "FORMULA_CONTRACT_MISSING", "无法识别公式任务契约。", undefined, "重新提交公式任务。", undefined, false));
      }
      const missingReads = getMissingRequiredReads(turn, contract);
      if (missingReads.length > 0) {
        return blocked(diagnostic(
          "input_read",
          "REQUIRED_RANGE_NOT_READ",
          "必须先读取数据源和用户已经指定的参考样例。",
          { missingRanges: missingReads.map((item) => item.raw) },
          "调用 range.read 读取缺失区域后，再判断场景和形状。",
          RANGE_READ_TOOL,
        ));
      }
      const activeCallIndex = findActiveToolCallIndex(turn, FORMULA_PREPARE_TOOL);
      const latestReadIndex = Math.max(...collectSuccessfulRangeReads(turn).map((read) => read.resultIndex));
      if (activeCallIndex >= 0 && latestReadIndex >= activeCallIndex) {
        return blocked(diagnostic(
          "prepare",
          "PREPARATION_BEFORE_READ_RESULTS",
          "formula.prepare 与 range.read 在同一模型轮次生成，模型尚未看到读取结果，不能据此判断场景和形状。",
          { latestReadResultIndex: latestReadIndex, prepareCallIndex: activeCallIndex },
          "等待读取结果返回后，在下一轮根据真实结构调用 formula.prepare。",
          FORMULA_PREPARE_TOOL,
        ));
      }
      const result = await original.execute(args, context);
      if (!result.success || !isPreparation(result.data)) return result;
      if (result.data.status === "ready" && contract.referenceMode === "none") {
        const executableChecks = result.data.acceptanceChecks.filter((check) => check.type !== "no_excel_error");
        if (executableChecks.length === 0) {
          return blocked(diagnostic(
            "prepare",
            "NO_SAMPLE_ACCEPTANCE_INSUFFICIENT",
            "参考样例为空时，不能只检查公式错误值，必须根据口语需求提供至少一项结构或业务不变量检查。",
            { supportedChecks: ["shape", "unique_key", "row_count", "aggregate_reconciliation", "sort_order", "lookup_consistency", "pattern_match", "boundary", "spot_check"] },
            "补充能够验证当前需求的 acceptanceChecks；如果业务口径存在多种解释，改用 needs_clarification。",
            FORMULA_PREPARE_TOOL,
          ));
        }
      }
      return result;
    },
  });
}

function guardFormulaRangeWrite(executors: Map<string, ToolExecutor>, turn: Turn): void {
  const original = executors.get(RANGE_WRITE_TOOL);
  if (!original) return;
  executors.set(RANGE_WRITE_TOOL, {
    name: RANGE_WRITE_TOOL,
    async execute(args, context) {
      const contract = parseFormulaTaskContract(turn);
      if (!contract) {
        return blocked(diagnostic("pre_write", "FORMULA_CONTRACT_MISSING", "无法识别公式任务契约。", undefined, "重新提交公式任务。", undefined, false));
      }
      const missingReads = getMissingRequiredReads(turn, contract);
      if (missingReads.length > 0) {
        return blocked(diagnostic(
          "input_read",
          "REQUIRED_RANGE_NOT_READ",
          "写入前仍有必需区域未读取。",
          { missingRanges: missingReads.map((item) => item.raw) },
          "先使用 range.read 读取缺失区域。",
          RANGE_READ_TOOL,
        ));
      }
      const preparation = getLatestFormulaPreparation(turn);
      if (!preparation) {
        return blocked(diagnostic("prepare", "FORMULA_PREPARATION_MISSING", "写入前必须提交结构化解题与验收契约。", undefined, "调用 formula.prepare。", FORMULA_PREPARE_TOOL));
      }
      if (preparation.preparation.status === "needs_clarification") {
        return blocked(diagnostic(
          "pre_write",
          "FORMULA_NEEDS_CLARIFICATION",
          "当前任务仍有会影响结果的实质歧义，不能写入。",
          { question: preparation.preparation.clarificationQuestion },
          "先向用户确认该问题，再在新任务中重新建立解题契约。",
          undefined,
          false,
        ));
      }
      const latestStructureReadIndex = Math.max(
        -1,
        ...collectSuccessfulRangeReads(turn)
          .filter((read) => read.call.arguments.expand !== "spill")
          .map((read) => read.resultIndex),
      );
      if (latestStructureReadIndex > preparation.resultIndex) {
        return blocked(diagnostic(
          "prepare",
          "FORMULA_PREPARATION_STALE",
          "解题契约之后又读取了新的数据结构，旧的场景、形状和验收判断已经失效。",
          { preparationResultIndex: preparation.resultIndex, latestStructureReadIndex },
          "根据最新读取结果重新调用 formula.prepare。",
          FORMULA_PREPARE_TOOL,
        ));
      }
      const methodology = findKnowledgeResult(turn, "formula_methodology", preparation.resultIndex, true);
      if (!methodology) {
        return blocked(diagnostic(
          "methodology",
          "METHODOLOGY_NOT_LOADED",
          "内置公式方法论未成功读取，不能写入。",
          undefined,
          "调用 knowledge.search，scope 使用 formula_methodology。",
          KNOWLEDGE_SEARCH_TOOL,
        ));
      }
      const activeCallIndex = findActiveToolCallIndex(turn, RANGE_WRITE_TOOL);
      if (activeCallIndex >= 0 && methodology.resultIndex >= activeCallIndex) {
        return blocked(diagnostic(
          "pre_write",
          "WRITE_BEFORE_METHODOLOGY_RESULT",
          "range.write 与方法论检索在同一模型轮次生成，模型尚未读取方法论结果。",
          undefined,
          "在下一模型轮次根据方法论生成并写入公式。",
          RANGE_WRITE_TOOL,
        ));
      }
      const staticError = validateFormulaWriteArgs(args, contract.targetRange);
      if (staticError) return blocked(staticError);
      return original.execute(args, context);
    },
  });
}

function guardFormulaVerification(executors: Map<string, ToolExecutor>, turn: Turn): void {
  if (!executors.has(FORMULA_VERIFY_TOOL)) return;
  executors.set(FORMULA_VERIFY_TOOL, {
    name: FORMULA_VERIFY_TOOL,
    async execute() {
      const report = verifyLatestFormulaWrite(turn);
      if (typeof report === "string") {
        return blocked(diagnostic("post_write", "FORMULA_VERIFICATION_NOT_READY", report, undefined, "先回读写入锚点的实际 spill 结果。", RANGE_READ_TOOL));
      }
      return { success: true, data: report };
    },
  });
}

function validateFormulaWriteArgs(
  args: Record<string, unknown>,
  expectedTarget: ReturnType<typeof parseFormulaRangeRef>,
): FormulaWorkflowDiagnostic | null {
  const sheetName = text(args.sheetName) || undefined;
  const range = text(args.range);
  const actualTarget = range ? parseFormulaRangeRef(`${sheetName ? `${sheetName}!` : ""}${range}`) : undefined;
  if (!actualTarget) {
    return diagnostic("pre_write", "INVALID_WRITE_TARGET", "range.write 缺少有效 sheetName 或 range。", { sheetName, range }, "补充有效目标锚点。", RANGE_WRITE_TOOL);
  }
  if (expectedTarget && !sameFormulaAnchor(actualTarget, expectedTarget)) {
    return diagnostic(
      "pre_write",
      "TARGET_RANGE_MISMATCH",
      "公式写入地址与用户指定目标锚点不一致。",
      { expected: expectedTarget.raw, actual: actualTarget.raw },
      "把 range.write 的 sheetName 和 range 改为用户指定锚点。",
      RANGE_WRITE_TOOL,
    );
  }
  if (!Array.isArray(args.values) || args.values.length === 0 || args.values.some((row) => !Array.isArray(row))) {
    return diagnostic("pre_write", "INVALID_FORMULA_VALUES", "range.write.values 必须是非空二维数组。", undefined, "使用二维 values 数组提交公式。", RANGE_WRITE_TOOL);
  }
  const formulas = (args.values as unknown[][]).flat().filter((value): value is string => typeof value === "string" && value.startsWith("="));
  if (formulas.length === 0) {
    return diagnostic("pre_write", "FORMULA_MISSING", "公式任务的 values 中没有以 = 开头的公式。", undefined, "生成 Excel/WPS 函数公式，不要写入手工结果值。", RANGE_WRITE_TOOL);
  }
  const dynamicFormulaCount = formulas.filter((formula) => /\b(?:FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|MAKEARRAY|MAP|BYROW|BYCOL|SCAN|REDUCE|GROUPBY|PIVOTBY)\s*\(/i.test(formula)).length;
  if (dynamicFormulaCount > 1) {
    return diagnostic(
      "pre_write",
      "MULTIPLE_DYNAMIC_ARRAY_ANCHORS",
      "检测到多个动态数组公式，动态数组应只写入一个锚点并自行溢出。",
      { dynamicFormulaCount },
      "仅保留目标锚点中的动态数组公式。",
      RANGE_WRITE_TOOL,
    );
  }
  return null;
}

function findKnowledgeResult(
  turn: Turn,
  scope: "formula_methodology" | "formula_scene",
  afterIndex: number,
  requireSuccess: boolean,
): { result: ToolResultItem; resultIndex: number } | null {
  const calls = new Map<string, ToolCallItem>();
  turn.items.forEach((item) => {
    if (item.type === "tool_call" && item.toolName === KNOWLEDGE_SEARCH_TOOL) calls.set(item.id, item);
  });
  for (let index = turn.items.length - 1; index > afterIndex; index--) {
    const item = turn.items[index];
    if (item.type !== "tool_result" || item.toolName !== KNOWLEDGE_SEARCH_TOOL) continue;
    const call = calls.get(item.toolCallId);
    if (call?.arguments.scope !== scope) continue;
    if (requireSuccess && item.isError) continue;
    return { result: item, resultIndex: index };
  }
  return null;
}

function hasSuccessfulMethodologyAfterPreparation(turn: Turn, preparationResultIndex: number): boolean {
  return Boolean(findKnowledgeResult(turn, "formula_methodology", preparationResultIndex, true));
}

function findActiveToolCallIndex(turn: Turn, toolName: string): number {
  for (let index = turn.items.length - 1; index >= 0; index--) {
    const item = turn.items[index];
    if (item.type === "tool_call" && item.toolName === toolName && (item.status === "running" || item.status === "pending")) {
      return index;
    }
  }
  return -1;
}

function buildReadShapeContext(turn: Turn): string {
  return collectSuccessfulRangeReads(turn).slice(-4).map((read) => {
    const rows = read.matrix?.length ?? 0;
    const columns = read.matrix?.reduce((max, row) => Math.max(max, row.length), 0) ?? 0;
    return `${read.ref.raw}=${rows}行x${columns}列`;
  }).join("；");
}

function formatPreparationForQuery(preparation: FormulaPreparation): string {
  return [
    `场景：${preparation.scenario}`,
    `输入形状：${preparation.inputShape}`,
    `输出形状：${preparation.outputShape}`,
    `业务键：${preparation.businessKeys.join("、") || "无"}`,
    `变换链：${preparation.transformChain.join(" -> ")}`,
    `约束：${preparation.constraints.join("；") || "无"}`,
    `验收：${preparation.acceptanceChecks.map((check) => check.description).join("；")}`,
  ].join("\n");
}

function diagnostic(
  stage: FormulaWorkflowDiagnostic["stage"],
  code: string,
  message: string,
  evidence: unknown,
  instruction: string,
  tool?: string,
  retryable = true,
): FormulaWorkflowDiagnostic {
  return {
    stage,
    code,
    message,
    evidence,
    nextAction: { tool, instruction },
    retryable,
  };
}

function blocked(value: FormulaWorkflowDiagnostic): { success: false; error: string } {
  return { success: false, error: formatFormulaDiagnostic(value) };
}

function isPreparation(value: unknown): value is FormulaPreparation {
  return typeof value === "object"
    && value !== null
    && "status" in value
    && ((value as { status?: unknown }).status === "ready" || (value as { status?: unknown }).status === "needs_clarification");
}

function text(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}
