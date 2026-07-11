import type { ToolExecutor, ToolResultItem, Turn } from "../../shared/types";
import {
  hasFormulaFunctionReference,
  isFormulaPromptContext,
} from "../../prompts/systemPrompt";

const KNOWLEDGE_SEARCH_TOOL = "knowledge.search";
const RANGE_READ_TOOL = "range.read";
const RANGE_WRITE_TOOL = "range.write";
const MAX_QUERY_CONTEXT_CHARS = 1_200;

export function isFormulaTurn(turn: Turn): boolean {
  return getTurnUserContent(turn).some((content) =>
    isFormulaPromptContext({ content }) || hasFormulaFunctionReference(content),
  );
}

export function hasFormulaKnowledgeSearchAttempt(turn: Turn): boolean {
  return turn.items.some(
    (item) => item.type === "tool_result" && item.toolName === KNOWLEDGE_SEARCH_TOOL,
  );
}

export function hasSuccessfulRangeRead(turn: Turn): boolean {
  return turn.items.some(
    (item) => item.type === "tool_result" && item.toolName === RANGE_READ_TOOL && !item.isError,
  );
}

export function shouldAutoSearchFormulaKnowledge(
  turn: Turn,
  executors: Map<string, ToolExecutor>,
): boolean {
  return isFormulaTurn(turn)
    && executors.has(KNOWLEDGE_SEARCH_TOOL)
    && hasSuccessfulRangeRead(turn)
    && !hasFormulaKnowledgeSearchAttempt(turn);
}

export function buildFormulaKnowledgeQuery(turn: Turn): string {
  const taskContent = getTurnUserContent(turn).join("\n").slice(0, MAX_QUERY_CONTEXT_CHARS);
  const readContext = turn.items
    .filter(
      (item): item is ToolResultItem =>
        item.type === "tool_result" && item.toolName === RANGE_READ_TOOL && !item.isError,
    )
    .slice(-3)
    .map((item) => stringifyCompact(item.result))
    .join("\n")
    .slice(0, MAX_QUERY_CONTEXT_CHARS);

  return [
    "Excel/WPS 动态数组公式方法论",
    taskContent ? `当前任务：${taskContent}` : "",
    readContext ? `已读取的数据上下文：${readContext}` : "",
    "只返回与当前任务结构、约束和验收直接相关的解题步骤与核心方法，忽略无关模式。",
  ].filter(Boolean).join("\n");
}

export function guardFormulaRangeWriteExecutors(
  executors: Map<string, ToolExecutor>,
  turn: Turn,
): Map<string, ToolExecutor> {
  const rangeWrite = executors.get(RANGE_WRITE_TOOL);
  if (
    !rangeWrite
    || !executors.has(KNOWLEDGE_SEARCH_TOOL)
    || !isFormulaTurn(turn)
  ) {
    return executors;
  }

  const guarded = new Map(executors);
  guarded.set(RANGE_WRITE_TOOL, {
    name: RANGE_WRITE_TOOL,
    async execute(args, context) {
      if (!hasFormulaKnowledgeSearchAttempt(turn)) {
        return {
          success: false,
          error: "公式写入前必须先读取真实数据并完成 knowledge.search 方法论检索。请先调用 range.read，再根据输入/输出形状和变换链检索知识库。",
        };
      }
      return rangeWrite.execute(args, context);
    },
  });
  return guarded;
}

function getTurnUserContent(turn: Turn): string[] {
  return turn.items
    .filter((item) => item.type === "user_message")
    .map((item) => item.content);
}

function stringifyCompact(value: unknown): string {
  if (typeof value === "string") return value.replace(/\s+/g, " ").trim();
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}
