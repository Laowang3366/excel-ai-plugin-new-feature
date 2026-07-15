import type { AgentTurnCallbacks } from "../../shared/types";
import {
  redactSensitiveText,
  redactSensitiveValue,
  summarizeValueForAudit,
} from "../../../shared/sensitiveData";

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
  const summary = summarizeValueForAudit(value);
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
    await appendToolExecutionLog({
      ...record,
      argumentsSummary: redactSensitiveText(record.argumentsSummary, 2000),
      resultSummary: redactSensitiveText(record.resultSummary, 2000),
      error: record.error ? redactSensitiveText(record.error, 2000) : undefined,
      metadata: record.metadata
        ? redactSensitiveValue(record.metadata) as Record<string, unknown>
        : undefined,
    });
  } catch (err: any) {
    callbacks.onEvent({
      type: "warning",
      message: `工具执行日志写入失败：${redactSensitiveText(err?.message || String(err), 1000)}`,
      threadId: record.threadId,
    });
  }
}
