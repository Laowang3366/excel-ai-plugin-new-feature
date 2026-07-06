import type { AgentTurnCallbacks } from "../../shared/types";

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

export function parseToolArguments(argsJson: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(argsJson || "{}");
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : { value: parsed };
  } catch {
    return { _raw: argsJson };
  }
}

export function summarizeForLog(value: unknown, maxLength = 2000): string {
  let summary: string;
  if (typeof value === "string") {
    summary = value;
  } else {
    try {
      summary = JSON.stringify(value) ?? String(value);
    } catch {
      summary = String(value);
    }
  }
  if (summary.length <= maxLength) return summary;
  return `${summary.slice(0, maxLength)}...`;
}

export async function logToolExecutionSafely(
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
