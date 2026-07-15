import type { RolloutItem, RolloutLine, ThreadId } from "../shared/types";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { clampNumber } from "../shared/numberLimits";
import {
  fieldAad,
  protectRequiredField,
  unprotectRequiredField,
} from "../../main-modules/localDataProtection/fieldCrypto";
import { extractRolloutSearchContent } from "./rolloutSearchContent";
import { getRolloutTurnId } from "./stateRuntimeMappers";
import type { RuntimeRolloutEvent, RuntimeRolloutSearchMatch } from "./stateRuntimeTypes";

type LogsWriteRunner = <T>(fn: () => T) => T;

function parseItemJson(rowId: string, itemJson: string): RolloutItem {
  const plain = unprotectRequiredField(
    itemJson,
    fieldAad("logs", "rollout_events", rowId, "item_json"),
  );
  return JSON.parse(plain) as RolloutItem;
}

function normalizeQueryTerms(query: string): string[] {
  return query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
}

function matchesAllTerms(content: string, terms: string[]): boolean {
  const normalized = content.toLocaleLowerCase();
  return terms.every((term) => normalized.includes(term));
}

function buildSnippet(content: string, terms: string[]): string {
  const normalized = content.toLocaleLowerCase();
  const firstIndex =
    terms
      .map((term) => normalized.indexOf(term))
      .filter((index) => index >= 0)
      .sort((a, b) => a - b)[0] ?? 0;
  const start = Math.max(0, firstIndex - 24);
  const end = Math.min(content.length, firstIndex + 80);
  return `${start > 0 ? "..." : ""}${content.slice(start, end)}${end < content.length ? "..." : ""}`;
}

export function appendRolloutItemsToLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
  items: RolloutItem[],
  runLogsWrite: LogsWriteRunner,
): void {
  if (items.length === 0) return;

  const nextIdStmt = logsDb.prepare(`SELECT COALESCE(MAX(id), 0) + 1 AS id FROM rollout_events`);
  const insert = logsDb.prepare(
    `INSERT INTO rollout_events (id, thread_id, turn_id, item_type, timestamp, item_json)
     VALUES (?, ?, ?, ?, ?, ?)`,
  );
  const insertSearch = logsDb.prepare(
    `INSERT INTO rollout_events_fts (rowid, thread_id, turn_id, item_type, content)
     VALUES (?, ?, ?, ?, ?)`,
  );
  const rows: RolloutLine[] = items.map((item) => ({
    timestamp: new Date().toISOString(),
    item,
  }));

  runLogsWrite(() => {
    for (const line of rows) {
      const itemJson = JSON.stringify(line.item);
      const turnId = getRolloutTurnId(line.item) ?? null;
      const nextId = Number((nextIdStmt.get() as { id: number }).id);
      const sealed = protectRequiredField(
        itemJson,
        fieldAad("logs", "rollout_events", String(nextId), "item_json"),
      );
      insert.run(nextId, threadId, turnId, line.item.type, line.timestamp, sealed);
      insertSearch.run(nextId, threadId, turnId, line.item.type, "");
    }
  });
}

export function listRolloutEventsFromLogs(
  logsDb: SqliteDatabase,
  threadId: ThreadId,
): RuntimeRolloutEvent[] {
  const rows = logsDb
    .prepare(`SELECT * FROM rollout_events WHERE thread_id = ? ORDER BY id ASC`)
    .all(threadId) as Record<string, any>[];

  return rows.map((row) => ({
    id: row.id,
    threadId: row.thread_id,
    turnId: row.turn_id ?? undefined,
    itemType: row.item_type,
    timestamp: row.timestamp,
    item: parseItemJson(String(row.id), row.item_json),
  }));
}

export function searchRolloutMatchesInLogs(
  logsDb: SqliteDatabase,
  query: string,
  options: { limit?: number } = {},
): RuntimeRolloutSearchMatch[] {
  const terms = normalizeQueryTerms(query);
  if (terms.length === 0) return [];

  const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
  const pageSize = 200;
  let offset = 0;
  const matches: RuntimeRolloutSearchMatch[] = [];

  while (matches.length < limit) {
    const rows = logsDb
      .prepare(
        `SELECT id, thread_id, turn_id, item_type, timestamp, item_json
         FROM rollout_events
         ORDER BY id DESC
         LIMIT ? OFFSET ?`,
      )
      .all(pageSize, offset) as Record<string, any>[];
    if (rows.length === 0) break;
    offset += rows.length;

    for (const row of rows) {
      try {
        const item = parseItemJson(String(row.id), row.item_json);
        const content = extractRolloutSearchContent(item);
        if (!matchesAllTerms(content, terms)) continue;
        matches.push({
          id: row.id,
          threadId: row.thread_id,
          turnId: row.turn_id ?? undefined,
          itemType: row.item_type,
          timestamp: row.timestamp,
          item,
          snippet: buildSnippet(content, terms),
        });
        if (matches.length >= limit) break;
      } catch {
        // skip corrupt rows
      }
    }
    if (rows.length < pageSize) break;
  }
  return matches;
}
