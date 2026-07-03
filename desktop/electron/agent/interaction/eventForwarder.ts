/**
 * Agent 事件转发与工具审批交互适配。
 *
 * 关联模块：
 * - core/agentLoop: 产生 AgentEvent 和流式增量回调。
 * - main.ts: 初始化 Agent 时注入 requestToolApproval。
 * - interaction/ipcAgentHandlers: 注册工具审批确认/取消 IPC。
 */

import { BrowserWindow, ipcMain } from "electron";
import type { AgentEvent } from "../shared/types";

/** 挂起的工具审批：toolCallId → { resolve, reject } */
export const pendingApprovals = new Map<string, {
  resolve: (approved: boolean, alwaysAllow?: boolean) => void;
  reject: (reason: string) => void;
}>();

export function createEventForwarder(mainWindowRef: () => BrowserWindow | null) {
  return {
    onEvent: (agentEvent: AgentEvent) => {
      const mw = mainWindowRef();
      if (mw && !mw.isDestroyed()) {
        mw.webContents.send("agent:event", agentEvent);
      }
    },
    onStreamDelta: (
      delta: string,
      itemType: "assistant_message" | "reasoning",
      roundId?: number,
      threadId?: string,
      clientId?: string
    ) => {
      const mw = mainWindowRef();
      if (mw && !mw.isDestroyed()) {
        mw.webContents.send("agent:streamDelta", { delta, itemType, roundId, threadId, clientId });
      }
    },
  };
}

const TOOL_APPROVAL_TIMEOUT_MS = 60_000;

export function requestToolApproval(
  mainWindowRef: () => BrowserWindow | null,
  params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
  }
): Promise<{ approved: boolean; alwaysAllow?: boolean }> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      if (pendingApprovals.has(params.toolCallId)) {
        pendingApprovals.delete(params.toolCallId);
        reject("审批超时（60秒无响应）");
      }
    }, TOOL_APPROVAL_TIMEOUT_MS);

    pendingApprovals.set(params.toolCallId, {
      resolve: (approved: boolean, alwaysAllow?: boolean) => {
        clearTimeout(timer);
        pendingApprovals.delete(params.toolCallId);
        resolve({ approved, alwaysAllow });
      },
      reject: (reason: string) => {
        clearTimeout(timer);
        pendingApprovals.delete(params.toolCallId);
        reject(reason);
      },
    });

    const mw = mainWindowRef();
    if (mw && !mw.isDestroyed()) {
      mw.webContents.send("agent:event", {
        type: "tool_approval_required",
        ...params,
      });
    } else {
      clearTimeout(timer);
      pendingApprovals.delete(params.toolCallId);
      reject("窗口已关闭");
    }
  });
}

export function registerToolApprovalHandlers(): void {
  ipcMain.handle("tool:confirm", async (_event, toolCallId: string, alwaysAllow?: boolean) => {
    const pending = pendingApprovals.get(toolCallId);
    if (pending) {
      pending.resolve(true, alwaysAllow);
    }
  });

  ipcMain.handle("tool:cancel", async (_event, toolCallId: string) => {
    const pending = pendingApprovals.get(toolCallId);
    if (pending) {
      pending.reject("用户取消了工具执行");
    }
  });
}
