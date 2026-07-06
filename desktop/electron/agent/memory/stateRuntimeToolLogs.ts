import type { ThreadId } from "../shared/types";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { clampNumber } from "../shared/numberLimits";
import { mapToolExecutionLog } from "./stateRuntimeMappers";
import type { RuntimeToolExecutionLogRecord } from "./stateRuntimeTypes";

export function appendToolExecutionLogToLogs(
  logsDb: SqliteDatabase,
  record: RuntimeToolExecutionLogRecord,
): void {
  logsDb.prepare(
    `INSERT INTO tool_execution_logs (
      thread_id, turn_id, tool_call_id, tool_name, status, duration_ms,
      timestamp, arguments_summary, result_summary, error, metadata_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(
    record.threadId,
    record.turnId,
    record.toolCallId,
    record.toolName,
    record.status,
    Math.max(0, Math.floor(record.durationMs)),
    record.timestamp,
    record.argumentsSummary,
    record.resultSummary,
    record.error ?? null,
    record.metadata ? JSON.stringify(record.metadata) : null,
  );
}

export function listToolExecutionLogsFromLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
  options: { limit?: number } = {},
): RuntimeToolExecutionLogRecord[] {
  const rows = logsDb.prepare(
    `SELECT * FROM tool_execution_logs
     WHERE thread_id = ?
     ORDER BY id ASC
     LIMIT ?`
  ).all(threadId, clampNumber(options.limit, { fallback: 200, min: 1, max: 1000 })) as Record<string, any>[];

  return rows.map(mapToolExecutionLog);
}
