import type { ThreadId, ThreadMetadata, ThreadRuntimeSnapshot } from "../shared/types";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { runSqliteTransaction } from "../storage/nodeSqlite";
import { mapThreadSnapshot } from "./stateRuntimeMappers";

export function upsertThreadSnapshotInDb(db: SqliteDatabase, metadata: ThreadMetadata): void {
  const upsert = db.prepare(
    `INSERT INTO thread_snapshots (
      thread_id, preview, name, model_provider, model, context_window_size,
      created_at, updated_at, active_turn_id, last_turn_status,
      total_token_usage, archived_at, folder_id, compacted_history
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(thread_id) DO UPDATE SET
      preview = excluded.preview,
      name = excluded.name,
      model_provider = excluded.model_provider,
      model = excluded.model,
      context_window_size = excluded.context_window_size,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at,
      active_turn_id = excluded.active_turn_id,
      last_turn_status = excluded.last_turn_status,
      total_token_usage = excluded.total_token_usage,
      archived_at = excluded.archived_at,
      folder_id = excluded.folder_id,
      compacted_history = excluded.compacted_history`
  );

  upsert.run(
    metadata.threadId,
    metadata.preview,
    metadata.name ?? null,
    metadata.modelProvider,
    metadata.model ?? null,
    metadata.contextWindowSize ?? null,
    metadata.createdAt,
    metadata.updatedAt,
    metadata.activeTurnId ?? null,
    metadata.lastTurnStatus ?? null,
    metadata.totalTokenUsage ? JSON.stringify(metadata.totalTokenUsage) : null,
    metadata.archivedAt ?? null,
    metadata.folderId ?? null,
    metadata.compactedHistory ? JSON.stringify(metadata.compactedHistory) : null
  );
}

export function getThreadSnapshotFromDb(db: SqliteDatabase, threadId: ThreadId): ThreadMetadata | null {
  const row = db.prepare(
    `SELECT * FROM thread_snapshots WHERE thread_id = ?`
  ).get(threadId) as Record<string, any> | undefined;
  return row ? mapThreadSnapshot(row) : null;
}

export function listThreadSnapshotsFromDb(db: SqliteDatabase): ThreadMetadata[] {
  const rows = db.prepare(
    `SELECT * FROM thread_snapshots ORDER BY updated_at DESC`
  ).all() as Record<string, any>[];
  return rows.map(mapThreadSnapshot);
}

export function deleteThreadStateFromDb(db: SqliteDatabase, threadId: ThreadId): void {
  runSqliteTransaction(db, () => {
    db.prepare("DELETE FROM thread_runtime WHERE thread_id = ?").run(threadId);
    db.prepare("DELETE FROM thread_names WHERE thread_id = ?").run(threadId);
    db.prepare("DELETE FROM thread_snapshots WHERE thread_id = ?").run(threadId);
  });
}

export function updateThreadRuntimeInDb(
  db: SqliteDatabase,
  snapshot: ThreadRuntimeSnapshot & { threadId: ThreadId }
): void {
  db.prepare(
    `INSERT INTO thread_runtime (
      thread_id, status, last_active_at, unloaded_at, idle_unload_ms, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(thread_id) DO UPDATE SET
      status = excluded.status,
      last_active_at = excluded.last_active_at,
      unloaded_at = excluded.unloaded_at,
      idle_unload_ms = excluded.idle_unload_ms,
      updated_at = excluded.updated_at`
  ).run(
    snapshot.threadId,
    snapshot.status,
    snapshot.lastActiveAt ?? null,
    snapshot.unloadedAt ?? null,
    snapshot.idleUnloadMs,
    Date.now()
  );
}

export function getThreadRuntimeFromDb(
  db: SqliteDatabase,
  threadId: ThreadId
): (ThreadRuntimeSnapshot & { threadId: ThreadId }) | null {
  const row = db.prepare(
    `SELECT * FROM thread_runtime WHERE thread_id = ?`
  ).get(threadId) as Record<string, any> | undefined;
  if (!row) return null;
  return {
    threadId: row.thread_id,
    status: row.status,
    lastActiveAt: row.last_active_at ?? undefined,
    unloadedAt: row.unloaded_at ?? undefined,
    idleUnloadMs: row.idle_unload_ms,
  };
}
