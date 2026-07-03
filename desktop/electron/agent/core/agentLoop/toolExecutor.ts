/**
 * 工具执行与审批管理
 *
 * 从 AgentLoop 中提取的工具执行相关逻辑：
 * - executeTool: 执行单个工具
 * - getToolDefinitions: 获取可用工具定义列表
 * - shouldRequireApproval: 判断工具是否需要审批
 * - requestApproval: 通过回调请求用户审批
 * - alwaysAllowTool: 标记工具为"始终允许"
 * - processToolCalls: 批量处理工具调用（执行+审批+事件）
 */

import * as os from "os";
import {
  type ToolCallItem,
  type ToolDefinition,
  type ToolExecutionContext,
  type ToolExecutionResult,
  type ToolExecutor,
  type Turn,
  type TurnItem,
  type AgentTurnCallbacks,
} from "../../shared/types";
import {
  TOOL_DEFINITIONS_MAP,
  ALL_TOOL_DEFINITIONS,
} from "../../tools/registry/toolDefinitions";
import { type ToolCallInfo } from "./streamCollector";
import { evaluateCommand, type CommandEvaluation } from "../../security/sandbox";
import { desanitizeToolName, sanitizeToolName } from "../../providers/openaiCompatibleClient";

// ============================================================
// 类型
// ============================================================

/** 工具审批配置 */
export interface ToolApprovalConfig {
  permissionMode: "normal" | "auto_approve_safe" | "confirm_all";
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    /** 沙箱策略给出的理由（命中规则 justification 合并）；UI 用于向用户解释为什么弹窗 */
    sandboxJustification?: string;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
}

export interface ToolExecutionLogRecord {
  threadId: string;
  turnId: string;
  toolCallId: string;
  toolName: string;
  status: "success" | "error" | "cancelled" | "blocked";
  durationMs: number;
  timestamp: number;
  argumentsSummary: string;
  resultSummary: string;
  error?: string;
  metadata?: Record<string, unknown>;
}

// ============================================================
// 单工具执行
// ============================================================

/**
 * 执行单个工具
 *
 * @param name - 工具名称
 * @param argsJson - JSON 格式的参数字符串
 * @param executors - 可用的工具执行器映射
 * @returns 执行结果
 */
export async function executeTool(
  name: string,
  argsJson: string,
  executors?: Map<string, ToolExecutor>,
  context?: ToolExecutionContext
): Promise<ToolExecutionResult> {
  const resolvedName = resolveExecutableToolName(name, executors);
  if (!executors || !resolvedName) {
    return {
      success: false,
      error: `未知工具: ${name}`,
    };
  }

  const executor = executors.get(resolvedName)!;
  try {
    const args = JSON.parse(argsJson || "{}");
    return await executor.execute(args, context);
  } catch (err: any) {
    return {
      success: false,
      error: `工具执行错误: ${err.message}`,
    };
  }
}

/**
 * 获取可用工具定义列表
 *
 * 只返回有对应 executor 的工具定义。
 */
export function getToolDefinitions(
  executors?: Map<string, ToolExecutor>
): ToolDefinition[] {
  if (!executors || executors.size === 0) return [];
  return ALL_TOOL_DEFINITIONS.filter((def) => executors.has(def.name));
}

// ============================================================
// 审批逻辑
// ============================================================

/** 始终允许的工具集合（模块级维护） */
const alwaysAllowedTools = new Set<string>();

/**
 * 判断工具是否需要审批
 *
 * 三种权限模式：
 * - normal（逐次确认）：所有工具调用都需要用户审批，无例外
 * - auto_approve_safe（自动审批）：safe 级别自动批准；moderate + dangerous 需要审批
 * - confirm_all（全部确认）：所有工具均自动批准，不要求手动审批。
 *   单元格清除、行列删除、子表操作、控件/图表删除等均直接放行。
 *   安全隐患由沙箱策略前置拦截（forbidden 直接拒绝）。
 *
 * alwaysAllowedTools 中的工具在三种模式下均跳过审批。
 */
export function shouldRequireApproval(
  toolName: string,
  permissionMode: "normal" | "auto_approve_safe" | "confirm_all" = "normal"
): boolean {
  switch (permissionMode) {
    case "normal":
      // 逐次确认：但始终允许的工具可跳过
      if (alwaysAllowedTools.has(toolName)) return false;
      return true;

    case "auto_approve_safe":
      // 自动审批：safe 级别自动批准，其余需要审批
      if (alwaysAllowedTools.has(toolName)) return false;
      const safeDef = TOOL_DEFINITIONS_MAP.get(toolName);
      return safeDef ? safeDef.riskLevel !== "safe" : true;

    case "confirm_all":
      // 全部确认：所有工具均自动批准，不要求手动审批。
      // 表格的清除、行列删除、子表操作、控件/图表删除等均直接放行。
      // 安全隐患由沙箱策略前置拦截（forbidden 直接拒绝）。
      return false;

    default:
      return true;
  }
}

/**
 * 请求工具审批
 */
export async function requestToolApproval(
  params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    /** 沙箱策略理由 */
    sandboxJustification?: string;
  },
  config: ToolApprovalConfig
): Promise<{ approved: boolean; alwaysAllow?: boolean }> {
  if (config.requestToolApproval) {
    return config.requestToolApproval(params);
  }
  return { approved: true };
}

/** 将工具标记为"始终允许" */
export function markToolAlwaysAllowed(toolName: string): void {
  alwaysAllowedTools.add(toolName);
}

/** 获取始终允许的工具集合（只读） */
export function getAlwaysAllowedTools(): ReadonlySet<string> {
  return alwaysAllowedTools;
}

/** 清除始终允许的工具集合（用于测试或重置） */
export function clearAlwaysAllowedTools(): void {
  alwaysAllowedTools.clear();
}

// ============================================================
// 批量工具处理
// ============================================================

/**
 * 处理一组工具调用：执行 + 审批 + 状态事件
 *
 * @param toolCalls - 待处理的工具调用列表
 * @param pendingToolCallItems - 流式阶段已创建的 tool_call items
 * @param turn - 当前 Turn
 * @param executors - 工具执行器映射
 * @param approvalConfig - 审批配置
 * @param callbacks - 事件回调
 * @param sessionStoreAppend - 持久化回调（调用方提供）
 */
export async function processToolCalls(
  toolCalls: ToolCallInfo[],
  pendingToolCallItems: Map<string, ToolCallItem>,
  turn: Turn,
  executors: Map<string, ToolExecutor>,
  approvalConfig: ToolApprovalConfig,
  callbacks: AgentTurnCallbacks,
  sessionStoreAppend: (threadId: string, turnId: string, item: TurnItem) => Promise<void>,
  appendToolExecutionLog?: (record: ToolExecutionLogRecord) => Promise<void>
): Promise<void> {
  for (const tc of toolCalls) {
    const startedAt = Date.now();
    const resolvedToolName = resolveExecutableToolName(tc.name, executors) ?? desanitizeToolName(tc.name);
    const toolDef = TOOL_DEFINITIONS_MAP.get(resolvedToolName) ?? TOOL_DEFINITIONS_MAP.get(tc.name);
    const canonicalToolName = toolDef?.name ?? resolvedToolName;

    // ============================================================
    // 沙箱策略前置评估（仅 shell.execute）
    // ============================================================
    // 对应 Codex execpolicy：在执行前先评估命令，forbidden 拒绝，prompt 强制审批。
    // 沙箱决策覆盖 permissionMode 与 alwaysAllowedTools —— forbidden 永拒，prompt 永审。
    let sandboxEvaluation: CommandEvaluation | null = null;
    let sandboxJustification: string | undefined;
    let sandboxForcedForbidden = false;
    let sandboxForcedApproval = false;

    if (canonicalToolName === "shell.execute") {
      try {
        const parsedArgs: Record<string, unknown> = (() => {
          try { return JSON.parse(tc.arguments || "{}"); } catch { return { _raw: tc.arguments }; }
        })();
        const cmd = (parsedArgs.command as string) || "";
        const workdir = (parsedArgs.workdir as string) || os.homedir();
        sandboxEvaluation = await evaluateCommand(cmd, workdir);

        if (sandboxEvaluation.decision === "forbidden") {
          sandboxForcedForbidden = true;
          sandboxJustification = sandboxEvaluation.violationMessage;
        } else if (sandboxEvaluation.decision === "prompt") {
          sandboxForcedApproval = true;
          sandboxJustification = sandboxEvaluation.evaluation.hits
            .filter((h) => h.rule.decision === "prompt")
            .map((h) => h.rule.justification || h.matchedPrefix.join(" "))
            .filter(Boolean)
            .join("；")
            || "命中安全策略 prompt 规则，需要用户确认";
        }
      } catch {
        // 评估自身异常时保守地进入审批流程，不直接放行
        sandboxForcedApproval = true;
        sandboxJustification = "命令策略评估异常，需要用户确认";
      }
    }

    // forbidden：直接拒绝，不进审批、不进 spawn
    if (sandboxForcedForbidden) {
      let item = pendingToolCallItems.get(tc.id);
      if (item && item.toolName !== canonicalToolName) {
        item.toolName = canonicalToolName;
      }
      if (!item) {
        try {
          const parsedArgs = JSON.parse(tc.arguments || "{}");
          item = {
            type: "tool_call", id: tc.id, toolName: canonicalToolName,
            arguments: parsedArgs, status: "failed", timestamp: Date.now(),
          };
          turn.items.push(item);
          await sessionStoreAppend(turn.threadId, turn.turnId, item);
          callbacks.onEvent({ type: "item_started", item });
        } catch {
          item = undefined;
        }
      }
      if (item) {
        item.status = "failed";
        callbacks.onEvent({ type: "item_updated", item });
      }
      const resultItem: TurnItem = {
        type: "tool_result",
        id: `result-${Date.now()}`,
        toolCallId: tc.id,
        toolName: canonicalToolName,
        result: sandboxJustification || "命令被安全策略拒绝",
        isError: true,
        timestamp: Date.now(),
      };
      turn.items.push(resultItem);
      await sessionStoreAppend(turn.threadId, turn.turnId, resultItem);
      callbacks.onEvent({ type: "item_started", item: resultItem });
      callbacks.onEvent({ type: "item_completed", item: resultItem });
      await logToolExecutionSafely(appendToolExecutionLog, {
        threadId: turn.threadId,
        turnId: turn.turnId,
        toolCallId: tc.id,
        toolName: canonicalToolName,
        status: "blocked",
        durationMs: Date.now() - startedAt,
        timestamp: Date.now(),
        argumentsSummary: summarizeForLog(parseToolArguments(tc.arguments)),
        resultSummary: summarizeForLog(resultItem.result),
        error: sandboxJustification || "命令被安全策略拒绝",
        metadata: {
          permissionMode: approvalConfig.permissionMode,
          sandboxDecision: sandboxEvaluation?.decision ?? "forbidden",
          riskLevel: toolDef?.riskLevel ?? "moderate",
        },
      }, callbacks);
      continue;
    }

    // 计算是否需要审批：策略 prompt 始终审批，覆盖 alwaysAllowedTools 与 permissionMode
    const needsApproval = sandboxForcedApproval
      ? true
      : shouldRequireApproval(canonicalToolName, approvalConfig.permissionMode);

    // 获取或创建 tool_call item
    let activeItem = pendingToolCallItems.get(tc.id);
    if (activeItem && activeItem.toolName !== canonicalToolName) {
      activeItem.toolName = canonicalToolName;
    }
    if (!activeItem) {
      // 防御性兜底：如果流式阶段未创建
      let parsedArgs: Record<string, unknown> = {};
      try {
        parsedArgs = JSON.parse(tc.arguments || "{}");
      } catch {
        parsedArgs = { _raw: tc.arguments };
      }
      const fallbackItem: ToolCallItem = {
        type: "tool_call",
        id: tc.id,
        toolName: canonicalToolName,
        arguments: parsedArgs,
        status: needsApproval ? "pending" : "running",
        timestamp: Date.now(),
      };
      turn.items.push(fallbackItem);
      await sessionStoreAppend(turn.threadId, turn.turnId, fallbackItem);
      callbacks.onEvent({ type: "item_started", item: fallbackItem });
      activeItem = fallbackItem;
    } else {
      // 更新状态
      if (!needsApproval) {
        activeItem.status = "running";
        callbacks.onEvent({ type: "item_updated", item: activeItem });
      }
    }

    // 审批流程
    if (needsApproval) {
      try {
        const approvalArgs = (activeItem.arguments && typeof activeItem.arguments === "object" && !("_raw" in (activeItem.arguments as Record<string, unknown>)))
          ? activeItem.arguments as Record<string, unknown>
          : (() => { try { return JSON.parse(tc.arguments || "{}"); } catch { return { _raw: tc.arguments }; } })();

        const approval = await requestToolApproval(
          {
            toolCallId: tc.id,
            toolName: canonicalToolName,
            arguments: approvalArgs,
            riskLevel: toolDef?.riskLevel || "moderate",
            description: toolDef?.description,
            sandboxJustification,
          },
          approvalConfig
        );

        if (!approval.approved) {
          activeItem.status = "failed";
          const resultItem: TurnItem = {
            type: "tool_result",
            id: `result-${Date.now()}`,
            toolCallId: tc.id,
            toolName: canonicalToolName,
            result: "用户取消了工具执行",
            isError: true,
            timestamp: Date.now(),
          };
          turn.items.push(resultItem);
          await sessionStoreAppend(turn.threadId, turn.turnId, resultItem);
          callbacks.onEvent({ type: "item_updated", item: activeItem });
          callbacks.onEvent({ type: "item_started", item: resultItem });
          callbacks.onEvent({ type: "item_completed", item: resultItem });
          await logToolExecutionSafely(appendToolExecutionLog, {
            threadId: turn.threadId,
            turnId: turn.turnId,
            toolCallId: tc.id,
            toolName: canonicalToolName,
            status: "cancelled",
            durationMs: Date.now() - startedAt,
            timestamp: Date.now(),
            argumentsSummary: summarizeForLog(approvalArgs),
            resultSummary: summarizeForLog(resultItem.result),
            error: "用户取消了工具执行",
            metadata: {
              permissionMode: approvalConfig.permissionMode,
              sandboxDecision: sandboxEvaluation?.decision,
              riskLevel: toolDef?.riskLevel ?? "moderate",
            },
          }, callbacks);
          continue;
        }

        if (approval.alwaysAllow) {
          markToolAlwaysAllowed(canonicalToolName);
        }

        activeItem.status = "running";
        callbacks.onEvent({ type: "item_updated", item: activeItem });
      } catch (err: any) {
        activeItem.status = "failed";
        const resultItem: TurnItem = {
          type: "tool_result",
          id: `result-${Date.now()}`,
          toolCallId: tc.id,
          toolName: canonicalToolName,
          result: err.message || "审批过程出错",
          isError: true,
          timestamp: Date.now(),
        };
        turn.items.push(resultItem);
        await sessionStoreAppend(turn.threadId, turn.turnId, resultItem);
        callbacks.onEvent({ type: "item_updated", item: activeItem });
        callbacks.onEvent({ type: "item_started", item: resultItem });
        callbacks.onEvent({ type: "item_completed", item: resultItem });
        await logToolExecutionSafely(appendToolExecutionLog, {
          threadId: turn.threadId,
          turnId: turn.turnId,
          toolCallId: tc.id,
          toolName: canonicalToolName,
          status: "error",
          durationMs: Date.now() - startedAt,
          timestamp: Date.now(),
          argumentsSummary: summarizeForLog(parseToolArguments(tc.arguments)),
          resultSummary: summarizeForLog(resultItem.result),
          error: err.message || "审批过程出错",
          metadata: {
            permissionMode: approvalConfig.permissionMode,
            sandboxDecision: sandboxEvaluation?.decision,
            riskLevel: toolDef?.riskLevel ?? "moderate",
            phase: "approval",
          },
        }, callbacks);
        continue;
      }
    }

    // 执行工具
    const result = await executeTool(
      canonicalToolName,
      tc.arguments,
      executors,
      sandboxEvaluation ? { sandboxEvaluation } : undefined
    );

    // 更新工具调用状态
    activeItem.status = result.success ? "completed" : "failed";
    callbacks.onEvent({ type: "item_updated", item: activeItem });

    // 记录工具结果
    const resultItem: TurnItem = {
      type: "tool_result",
      id: `result-${Date.now()}`,
      toolCallId: tc.id,
      toolName: canonicalToolName,
      result: result.data || result.error,
      isError: !result.success,
      timestamp: Date.now(),
    };
    turn.items.push(resultItem);
    await sessionStoreAppend(turn.threadId, turn.turnId, resultItem);
    callbacks.onEvent({ type: "item_completed", item: activeItem });
    callbacks.onEvent({ type: "item_started", item: resultItem });
    callbacks.onEvent({ type: "item_completed", item: resultItem });
    await logToolExecutionSafely(appendToolExecutionLog, {
      threadId: turn.threadId,
      turnId: turn.turnId,
      toolCallId: tc.id,
      toolName: canonicalToolName,
      status: result.success ? "success" : "error",
      durationMs: Date.now() - startedAt,
      timestamp: Date.now(),
      argumentsSummary: summarizeForLog(activeItem.arguments),
      resultSummary: summarizeForLog(resultItem.result),
      error: result.success ? undefined : result.error,
      metadata: {
        permissionMode: approvalConfig.permissionMode,
        sandboxDecision: sandboxEvaluation?.decision,
        riskLevel: toolDef?.riskLevel ?? "moderate",
      },
    }, callbacks);
  }
}

function resolveExecutableToolName(
  name: string,
  executors?: Map<string, ToolExecutor>
): string | null {
  const desanitized = desanitizeToolName(name);
  const candidates = Array.from(new Set([
    name,
    desanitized,
    sanitizeToolName(name),
    sanitizeToolName(desanitized),
    name.replace(/\.(?=[^.]+$)/, "_"),
    desanitized.replace(/\.(?=[^.]+$)/, "_"),
  ]));
  for (const candidate of candidates) {
    if (executors?.has(candidate)) return candidate;
  }
  return null;
}

function parseToolArguments(argsJson: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(argsJson || "{}");
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : { value: parsed };
  } catch {
    return { _raw: argsJson };
  }
}

function summarizeForLog(value: unknown, maxLength = 2000): string {
  let summary: string;
  if (typeof value === "string") {
    summary = value;
  } else {
    try {
      summary = JSON.stringify(value);
    } catch {
      summary = String(value);
    }
  }
  if (summary.length <= maxLength) return summary;
  return `${summary.slice(0, maxLength)}...`;
}

async function logToolExecutionSafely(
  appendToolExecutionLog: ((record: ToolExecutionLogRecord) => Promise<void>) | undefined,
  record: ToolExecutionLogRecord,
  callbacks: AgentTurnCallbacks
): Promise<void> {
  if (!appendToolExecutionLog) return;
  try {
    await appendToolExecutionLog(record);
  } catch (err: any) {
    callbacks.onEvent({
      type: "warning",
      message: `工具执行日志写入失败：${err?.message || String(err)}`,
      threadId: record.threadId,
    });
  }
}
