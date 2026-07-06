import type { RolloutItem, RolloutLine, ThreadId } from "../shared/types";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { clampNumber } from "../shared/numberLimits";
import { extractRolloutSearchContent } from "./rolloutSearchContent";
import { buildRolloutFtsQuery, getRolloutTurnId } from "./stateRuntimeMappers";
import type {
  RuntimeRolloutEvent,
  RuntimeRolloutSearchMatch,
} from "./stateRuntimeTypes";

type LogsWriteRunner = <T>(fn: () => T) => T;

export function appendRolloutItemsToLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
  items: RolloutItem[],
  runLogsWrite: LogsWriteRunner,
): void {
  if (items.length === 0) return;

  const insert = logsDb.prepare(
    `INSERT INTO rollout_events (thread_id, turn_id, item_type, timestamp, item_json)
     VALUES (?, ?, ?, ?, ?)`
  );
  const insertSearch = logsDb.prepare(
    `INSERT INTO rollout_events_fts (rowid, thread_id, turn_id, item_type, content, item_json)
     VALUES (?, ?, ?, ?, ?, ?)`
  );
  const rows: RolloutLine[] = items.map((item) => ({
    timestamp: new Date().toISOString(),
    item,
  }));

  runLogsWrite(() => {
    for (const line of rows) {
      const itemJson = JSON.stringify(line.item);
      const turnId = getRolloutTurnId(line.item) ?? null;
      const result = insert.run(
        threadId,
        turnId,
        line.item.type,
        line.timestamp,
        itemJson,
      );
      insertSearch.run(
        Number(result.lastInsertRowid),
        threadId,
        turnId,
        line.item.type,
        extractRolloutSearchContent(line.item),
        itemJson,
      );
    }
  });
}

export function listRolloutEventsFromLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
): RuntimeRolloutEvent[] {
  const rows = logsDb.prepare(
    `SELECT * FROM rollout_events WHERE thread_id = ? ORDER BY id ASC`
  ).all(threadId) as Record<string, any>[];

  return rows.map((row) => ({
    id: row.id,
    threadId: row.thread_id,
    turnId: row.turn_id ?? undefined,
    itemType: row.item_type,
    timestamp: row.timestamp,
    item: JSON.parse(row.item_json),
  }));
}

export function searchRolloutMatchesInLogs(
  logsDb: SqliteDatabase,
  query: string,
  options: { limit?: number } = {},
): RuntimeRolloutSearchMatch[] {
  const normalized = query.trim();
  if (!normalized) return [];

  const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
  const rows = logsDb.prepare(
    `SELECT
        e.id,
        e.thread_id,
        e.turn_id,
        e.item_type,
        e.timestamp,
        e.item_json,
        snippet(rollout_events_fts, 3, '[', ']', '...', 16) AS snippet
      FROM rollout_events_fts
      JOIN rollout_events e ON e.id = rollout_events_fts.rowid
      WHERE rollout_events_fts.content MATCH ?
      ORDER BY e.id DESC
      LIMIT ?`
  ).all(buildRolloutFtsQuery(normalized), limit) as Record<string, any>[];

  return rows.map((row) => ({
    id: row.id,
    threadId: row.thread_id,
    turnId: row.turn_id ?? undefined,
    itemType: row.item_type,
    timestamp: row.timestamp,
    item: JSON.parse(row.item_json),
    snippet: row.snippet ?? "",
  }));
}
