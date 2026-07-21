import { useCallback, useState } from "react";
import type { AgentContentPart } from "@shared/agent";
import type { ChatToolExecuteResult } from "@shared/agentChat";
import { useChatSession } from "../../chat/ChatSessionContext";

export interface TaskSubmitOutcome {
  accepted: boolean;
  assistantText?: string;
  turnStatus?: string;
}

export function useTaskSubmit() {
  const { send, executeTool, view, adapter, approve, reject } = useChatSession();
  const [error, setError] = useState<string | undefined>();
  const [lastResult, setLastResult] = useState<string | undefined>();
  const busy =
    view.status === "running" ||
    view.status === "awaiting_approval" ||
    view.status === "stopping";

  const submit = useCallback(
    async (
      payload: string,
      options?: { contentParts?: AgentContentPart[] },
    ): Promise<TaskSubmitOutcome> => {
      setError(undefined);
      setLastResult(undefined);
      if (!adapter) {
        setError("宿主未就绪，无法提交任务");
        return { accepted: false };
      }
      if (!view.canSend) {
        setError("聊天忙或正在等待批准，请稍后再提交");
        return { accepted: false };
      }
      const outcome = await send(payload, options);
      if (!outcome.accepted) {
        setError(
          outcome.turnStatus === "busy"
            ? "会话忙，请等待当前任务完成"
            : outcome.turnStatus === "preflight_failed"
              ? "提交前检查失败（请检查模型供应商配置）"
              : "未能提交任务",
        );
        return {
          accepted: false,
          turnStatus: outcome.turnStatus,
          assistantText: outcome.assistantText,
        };
      }
      setLastResult(
        outcome.turnStatus === "completed"
          ? "任务完成（详见下方预览/聊天）"
          : view.pendingApproval
            ? "已提交，等待批准（见聊天或本页审批）"
            : "已提交到聊天会话",
      );
      return {
        accepted: true,
        turnStatus: outcome.turnStatus,
        assistantText: outcome.assistantText,
      };
    },
    [adapter, send, view.canSend, view.pendingApproval],
  );

  const runTool = useCallback(
    async (
      toolName: string,
      args: Record<string, unknown>,
      options?: { toolCallId?: string },
    ): Promise<ChatToolExecuteResult> => {
      setError(undefined);
      if (!adapter) {
        const r = { ok: false, tool: toolName, error: "宿主未就绪" };
        setError(r.error);
        return r;
      }
      if (busy) {
        const r = { ok: false, tool: toolName, error: "会话忙，请稍后再试" };
        setError(r.error);
        return r;
      }
      const result = await executeTool(toolName, args, options);
      if (!result.ok) {
        setError(result.error || "工具执行失败");
      } else {
        setLastResult("写入已完成（经审批边界）");
      }
      return result;
    },
    [adapter, busy, executeTool],
  );

  return {
    submit,
    runTool,
    busy,
    error,
    lastResult,
    setError,
    setLastResult,
    view,
    adapter,
    approve,
    reject,
  };
}
