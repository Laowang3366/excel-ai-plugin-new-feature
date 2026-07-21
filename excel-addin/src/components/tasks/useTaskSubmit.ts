import { useCallback, useState } from "react";
import type { AgentContentPart } from "@shared/agent";
import { useChatSession } from "../../chat/ChatSessionContext";

export function useTaskSubmit() {
  const { send, view, adapter } = useChatSession();
  const [error, setError] = useState<string | undefined>();
  const [lastResult, setLastResult] = useState<string | undefined>();
  const busy = view.status === "running" || view.status === "awaiting_approval" || view.status === "stopping";

  const submit = useCallback(
    async (
      payload: string,
      options?: { contentParts?: AgentContentPart[] },
    ): Promise<boolean> => {
      setError(undefined);
      setLastResult(undefined);
      if (!adapter) {
        setError("宿主未就绪，无法提交任务");
        return false;
      }
      if (!view.canSend) {
        setError("聊天忙或正在等待批准，请稍后再提交");
        return false;
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
        return false;
      }
      setLastResult(
        outcome.turnStatus === "completed"
          ? "任务完成（详见聊天）"
          : view.pendingApproval
            ? "已提交，等待批准（见聊天）"
            : "已提交到聊天会话",
      );
      return true;
    },
    [adapter, send, view.canSend, view.pendingApproval],
  );

  return { submit, busy, error, lastResult, setError, view, adapter };
}
