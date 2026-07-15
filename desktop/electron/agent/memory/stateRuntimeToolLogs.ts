import type { ThreadId } from "../shared/types";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { clampNumber } from "../shared/numberLimits";
import {
  fieldAad,
  protectFieldValue,
  protectRequiredField,
} from "../../main-modules/localDataProtection/fieldCrypto";
import { mapToolExecutionLog } from "./stateRuntimeMappers";
import type { RuntimeToolExecutionLogRecord } from "./stateRuntimeTypes";

export function appendToolExecutionLogToLogs(
  logsDb: SqliteDatabase,
  record: RuntimeToolExecutionLogRecord,
  runLogsWrite: <T>(fn: () => T) => T = (fn) => fn(),
): void {
  runLogsWrite(() => {
    const nextIdRow = logsDb
      .prepare(`SELECT COALESCE(MAX(id), 0) + 1 AS id FROM tool_execution_logs`)
      .get() as { id: number };
    const rowId = String(nextIdRow.id);
    logsDb
      .prepare(
        `INSERT INTO tool_execution_logs (
          id, thread_id, turn_id, tool_call_id, tool_name, status, duration_ms,
          timestamp, arguments_summary, result_summary, error, metadata_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(
        nextIdRow.id,
        record.threadId,
        record.turnId,
        record.toolCallId,
        record.toolName,
        record.status,
        Math.max(0, Math.floor(record.durationMs)),
        record.timestamp,
        protectRequiredField(
          record.argumentsSummary,
          fieldAad("logs", "tool_execution_logs", rowId, "arguments_summary"),
        ),
        protectRequiredField(
          record.resultSummary,
          fieldAad("logs", "tool_execution_logs", rowId, "result_summary"),
        ),
        protectFieldValue(
          record.error ?? null,
          fieldAad("logs", "tool_execution_logs", rowId, "error"),
        ),
        record.metadata
          ? protectRequiredField(
              JSON.stringify(record.metadata),
              fieldAad("logs", "tool_execution_logs", rowId, "metadata_json"),
            )
          : null,
      );
  });
}

export function listToolExecutionLogsFromLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
  options: { limit?: number } = {},
): RuntimeToolExecutionLogRecord[] {
  const rows = logsDb
    .prepare(
      `SELECT * FROM tool_execution_logs
       WHERE thread_id = ?
       ORDER BY id ASC
       LIMIT ?`,
    )
    .all(threadId, clampNumber(options.limit, { fallback: 200, min: 1, max: 1000 })) as Record<
    string,
    any
  >[];

  return rows.map(mapToolExecutionLog);
}
