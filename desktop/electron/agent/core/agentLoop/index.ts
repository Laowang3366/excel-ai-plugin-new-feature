/**
 * agentLoop 子模块 barrel re-export
 *
 * 向后兼容：外部消费者通过 "./agentLoop" 导入的符号不变。
 */

// 核心 AgentLoop 类
export { AgentLoop, type AgentLoopConfig } from "./agentLoop";

// 流式收集（ToolCallInfo 不导出，与 aiClientTypes 中的 ToolCallInfo 命名冲突）
export { collectStreamEvents, emitInterruptedProgress } from "./streamCollector";
export type { StreamParams, StreamResult } from "./streamCollector";

// 工具执行与审批
export {
  executeTool,
  getToolDefinitions,
  shouldRequireApproval,
  requestToolApproval,
  markToolAlwaysAllowed,
  getAlwaysAllowedTools,
  clearAlwaysAllowedTools,
  processToolCalls,
} from "./toolExecutor";
export type { ToolApprovalConfig } from "./toolExecutor";

// 压缩摘要与会话配置
export { generateSummary } from "./summaryGenerator";
export { buildSessionCompactionConfig } from "./sessionCompactionConfig";

// 流式参数构建
export {
  buildEffectiveSystemPrompt,
} from "./buildStreamParams";

// 输出预算与线程拆分模块
export { resolveMaxTokens } from "./maxTokens";
export { TurnState } from "./turnState";
export {
  DEFAULT_THREAD_IDLE_UNLOAD_MS,
  ThreadStateManager,
} from "./threadStateManager";
export type { ThreadRuntimeSnapshot, ThreadRuntimeStatus } from "../../shared/types";
export { createAgentThread, loadAgentThread } from "./threadLifecycle";
export { createTurn, createUserMessageItem, completeTurn } from "./turnRunner";
