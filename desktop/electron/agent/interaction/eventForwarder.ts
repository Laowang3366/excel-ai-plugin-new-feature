/**
 * Agent 事件转发与工具审批交互适配。
 *
 * 关联模块：
 * - core/agentLoop: 产生 AgentEvent 和流式增量回调。
 * - main.ts: 初始化 Agent 时注入 requestToolApproval。
 * - interaction/ipcAgentHandlers: 注册工具审批确认/取消 IPC。
 */

import { BrowserWindow, ipcMain } from "electron";
import {
  ToolCancelInput,
  ToolConfirmInput,
  validateInput,
} from "../../shared/ipcSchemas";
import type { AgentEvent } from "../shared/types";

export const STREAM_DELTA_FORWARD_FLUSH_MS = 32;

export type StreamDeltaPayload = {
  delta: string;
  itemType: "assistant_message" | "reasoning";
  roundId?: number;
  threadId?: string;
  clientId?: string;
};

/** 挂起的工具审批：toolCallId → { resolve, reject } */
export const pendingApprovals = new Map<string, {
  resolve: (approved: boolean, alwaysAllow?: boolean) => void;
  reject: (reason: string) => void;
}>();

export function mergeStreamDeltas(deltas: StreamDeltaPayload[]): StreamDeltaPayload[] {
  const merged: StreamDeltaPayload[] = [];
  for (const delta of deltas) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.itemType === delta.itemType &&
      previous.roundId === delta.roundId &&
      previous.threadId === delta.threadId &&
      previous.clientId === delta.clientId
    ) {
      previous.delta += delta.delta;
    } else {
      merged.push({ ...delta });
    }
  }
  return merged;
}

export function createEventForwarder(mainWindowRef: () => BrowserWindow | null) {
  let pendingStreamDeltas: StreamDeltaPayload[] = [];
  let flushTimer: ReturnType<typeof setTimeout> | null = null;

  const flushStreamDeltas = () => {
    if (flushTimer) {
      clearTimeout(flushTimer);
      flushTimer = null;
    }
    if (pendingStreamDeltas.length === 0) return;
    const batch = mergeStreamDeltas(pendingStreamDeltas);
    pendingStreamDeltas = [];

    const mw = mainWindowRef();
    if (!mw || mw.isDestroyed()) return;
    for (const payload of batch) {
      mw.webContents.send("agent:event", { type: "stream_delta", ...payload });
    }
  };

  const scheduleStreamDeltaFlush = () => {
    if (flushTimer) return;
    flushTimer = setTimeout(flushStreamDeltas, STREAM_DELTA_FORWARD_FLUSH_MS);
  };

  return {
    onEvent: (agentEvent: AgentEvent) => {
      flushStreamDeltas();
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
      pendingStreamDeltas.push({ delta, itemType, roundId, threadId, clientId });
      scheduleStreamDeltaFlush();
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
  ipcMain.handle("tool:confirm", async (_event, toolCallId: unknown, alwaysAllow?: unknown) => {
    const validated = validateInput(ToolConfirmInput, { toolCallId, alwaysAllow });
    const pending = pendingApprovals.get(validated.toolCallId);
    if (pending) {
      pending.resolve(true, validated.alwaysAllow);
    }
  });

  ipcMain.handle("tool:cancel", async (_event, toolCallId: unknown) => {
    const validated = validateInput(ToolCancelInput, toolCallId);
    const pending = pendingApprovals.get(validated);
    if (pending) {
      pending.reject("用户取消了工具执行");
    }
  });
}
