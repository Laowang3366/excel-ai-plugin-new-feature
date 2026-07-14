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

import {
  type ToolCallItem,
  type ToolDefinition,
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
import { desanitizeToolName } from "../../providers/openaiCompatibleClient";
import {
  logToolExecutionSafely,
  parseToolArguments,
  summarizeForLog,
  type ToolExecutionLogRecord,
} from "./toolExecutionLog";
import { resolveExecutableToolName } from "./toolNameResolution";
import { createToolResultItem } from "./toolResultItems";
import {
  markToolAlwaysAllowed,
  requestToolApproval,
  shouldRequireApproval,
  type ToolApprovalConfig,
} from "./toolApproval";

export type { ToolExecutionLogRecord } from "./toolExecutionLog";
export {
  clearAlwaysAllowedTools,
  getAlwaysAllowedTools,
  markToolAlwaysAllowed,
  requestToolApproval,
  shouldRequireApproval,
} from "./toolApproval";
export type { ToolApprovalConfig } from "./toolApproval";

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
  executors?: Map<string, ToolExecutor>
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
    return await executor.execute(args);
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
 * @param throwIfAborted - 在开始下一项工作前检查当前 Turn 是否已中断
 */
export async function processToolCalls(
  toolCalls: ToolCallInfo[],
  pendingToolCallItems: Map<string, ToolCallItem>,
  turn: Turn,
  executors: Map<string, ToolExecutor>,
  approvalConfig: ToolApprovalConfig,
  callbacks: AgentTurnCallbacks,
  sessionStoreAppend: (threadId: string, turnId: string, item: TurnItem) => Promise<void>,
  appendToolExecutionLog?: (record: ToolExecutionLogRecord) => Promise<void>,
  throwIfAborted?: () => void
): Promise<void> {
  for (const tc of toolCalls) {
    throwIfAborted?.();
    const startedAt = Date.now();
    const resolvedToolName = resolveExecutableToolName(tc.name, executors) ?? desanitizeToolName(tc.name);
    const toolDef = TOOL_DEFINITIONS_MAP.get(resolvedToolName) ?? TOOL_DEFINITIONS_MAP.get(tc.name);
    const canonicalToolName = toolDef?.name ?? resolvedToolName;

    const needsApproval = shouldRequireApproval(canonicalToolName, approvalConfig.permissionMode);

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
          },
          approvalConfig
        );

        if (!approval.approved) {
          activeItem.status = "failed";
          const resultItem: TurnItem = createToolResultItem({
            toolCallId: tc.id,
            toolName: canonicalToolName,
            result: "用户取消了工具执行",
            isError: true,
          });
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
        const resultItem: TurnItem = createToolResultItem({
          toolCallId: tc.id,
          toolName: canonicalToolName,
          result: err.message || "审批过程出错",
          isError: true,
        });
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
            riskLevel: toolDef?.riskLevel ?? "moderate",
            phase: "approval",
          },
        }, callbacks);
        continue;
      }
    }

    // 执行工具
    throwIfAborted?.();
    const result = await executeTool(canonicalToolName, tc.arguments, executors);

    // 更新工具调用状态
    activeItem.status = result.success ? "completed" : "failed";
    callbacks.onEvent({ type: "item_updated", item: activeItem });

    // 记录工具结果
    const resultItem: TurnItem = createToolResultItem({
      toolCallId: tc.id,
      toolName: canonicalToolName,
      result: result.success ? result.data : result.error,
      isError: !result.success,
    });
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
        riskLevel: toolDef?.riskLevel ?? "moderate",
      },
    }, callbacks);
  }
}
