import type { Database as BetterSqliteDatabase } from "better-sqlite3";

import type { RuntimeDbName, RuntimeMigration } from "./stateRuntimeTypes";

export const STATE_RUNTIME_MIGRATIONS: Record<RuntimeDbName, RuntimeMigration[]> = {
  state: [
    {
      id: "001_state_runtime",
      sql: `
        CREATE TABLE IF NOT EXISTS thread_snapshots (
          thread_id TEXT PRIMARY KEY,
          preview TEXT NOT NULL DEFAULT '',
          name TEXT,
          model_provider TEXT NOT NULL,
          model TEXT,
          context_window_size INTEGER,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL,
          active_turn_id TEXT,
          last_turn_status TEXT,
          total_token_usage TEXT,
          archived_at INTEGER,
          folder_id TEXT,
          compacted_history TEXT
        );

        CREATE TABLE IF NOT EXISTS thread_runtime (
          thread_id TEXT PRIMARY KEY,
          status TEXT NOT NULL,
          last_active_at INTEGER,
          unloaded_at INTEGER,
          idle_unload_ms INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_thread_snapshots_updated_at
          ON thread_snapshots(updated_at);
        CREATE INDEX IF NOT EXISTS idx_thread_runtime_status
          ON thread_runtime(status);
      `,
    },
    {
      id: "002_thread_names",
      sql: `
        CREATE TABLE IF NOT EXISTS thread_names (
          thread_id TEXT PRIMARY KEY,
          name TEXT NOT NULL,
          updated_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_thread_names_updated_at
          ON thread_names(updated_at);
      `,
    },
  ],
  logs: [
    {
      id: "001_rollout_logs",
      sql: `
        CREATE TABLE IF NOT EXISTS rollout_events (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          thread_id TEXT NOT NULL,
          turn_id TEXT,
          item_type TEXT NOT NULL,
          timestamp TEXT NOT NULL,
          item_json TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_rollout_events_thread_id_id
          ON rollout_events(thread_id, id);
        CREATE INDEX IF NOT EXISTS idx_rollout_events_item_type
          ON rollout_events(item_type);
      `,
    },
    {
      id: "002_rollout_search",
      sql: `
        CREATE VIRTUAL TABLE IF NOT EXISTS rollout_events_fts USING fts5(
          thread_id UNINDEXED,
          turn_id UNINDEXED,
          item_type UNINDEXED,
          content,
          item_json UNINDEXED
        );
      `,
    },
    {
      id: "003_tool_execution_logs",
      sql: `
        CREATE TABLE IF NOT EXISTS tool_execution_logs (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          thread_id TEXT NOT NULL,
          turn_id TEXT NOT NULL,
          tool_call_id TEXT NOT NULL,
          tool_name TEXT NOT NULL,
          status TEXT NOT NULL,
          duration_ms INTEGER NOT NULL,
          timestamp INTEGER NOT NULL,
          arguments_summary TEXT NOT NULL,
          result_summary TEXT NOT NULL,
          error TEXT,
          metadata_json TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_tool_execution_logs_thread_id_id
          ON tool_execution_logs(thread_id, id);
        CREATE INDEX IF NOT EXISTS idx_tool_execution_logs_tool_status
          ON tool_execution_logs(tool_name, status);
        CREATE INDEX IF NOT EXISTS idx_tool_execution_logs_timestamp
          ON tool_execution_logs(timestamp);
      `,
    },
    {
      id: "005_restore_rollout_fts5",
      sql: `
        DROP TABLE IF EXISTS rollout_events_fts;

        CREATE VIRTUAL TABLE IF NOT EXISTS rollout_events_fts USING fts5(
          thread_id UNINDEXED,
          turn_id UNINDEXED,
          item_type UNINDEXED,
          content,
          item_json UNINDEXED
        );
      `,
    },
  ],
  goals: [
    {
      id: "001_goals",
      sql: `
        CREATE TABLE IF NOT EXISTS goals (
          goal_id TEXT PRIMARY KEY,
          thread_id TEXT,
          objective TEXT NOT NULL,
          status TEXT NOT NULL,
          token_budget INTEGER,
          token_usage INTEGER,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL,
          completed_at INTEGER,
          payload_json TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_goals_thread_id
          ON goals(thread_id);
        CREATE INDEX IF NOT EXISTS idx_goals_status
          ON goals(status);
      `,
    },
  ],
  memories: [
    {
      id: "001_memories",
      sql: `
        CREATE TABLE IF NOT EXISTS memories (
          memory_id TEXT PRIMARY KEY,
          namespace TEXT NOT NULL,
          content TEXT NOT NULL,
          metadata_json TEXT,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_memories_namespace_updated_at
          ON memories(namespace, updated_at);
      `,
    },
    {
      id: "002_long_term_memories",
      sql: `
        CREATE TABLE IF NOT EXISTS long_term_memories (
          memory_id TEXT PRIMARY KEY,
          namespace TEXT NOT NULL,
          kind TEXT NOT NULL,
          visibility TEXT NOT NULL,
          status TEXT NOT NULL,
          content TEXT NOT NULL,
          summary TEXT,
          confidence REAL,
          source_thread_id TEXT,
          source_event_id INTEGER,
          workspace_fingerprint TEXT,
          expires_at INTEGER,
          metadata_json TEXT,
          citations_json TEXT,
          created_at INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS memory_pipeline_state (
          pipeline_id TEXT PRIMARY KEY,
          last_event_id INTEGER NOT NULL,
          updated_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_long_term_memories_visibility_status_updated_at
          ON long_term_memories(visibility, status, updated_at);
        CREATE INDEX IF NOT EXISTS idx_long_term_memories_namespace_kind_status
          ON long_term_memories(namespace, kind, status);
        CREATE INDEX IF NOT EXISTS idx_long_term_memories_source_thread
          ON long_term_memories(source_thread_id);
      `,
    },
  ],
};

export function configureRuntimeDatabase(db: BetterSqliteDatabase): void {
  db.pragma("journal_mode = WAL");
  db.pragma("foreign_keys = ON");
  db.pragma("busy_timeout = 5000");
}

export function applyRuntimeMigrations(
  db: BetterSqliteDatabase,
  migrations: RuntimeMigration[]
): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      id TEXT PRIMARY KEY,
      applied_at INTEGER NOT NULL
    );
  `);

  const applied = new Set(listAppliedRuntimeMigrations(db));
  const insert = db.prepare(
    `INSERT INTO schema_migrations (id, applied_at) VALUES (?, ?)`
  );

  const run = db.transaction(() => {
    for (const migration of migrations) {
      if (applied.has(migration.id)) continue;
      db.exec(migration.sql);
      insert.run(migration.id, Date.now());
    }
  });
  run();
}

export function listAppliedRuntimeMigrations(db: BetterSqliteDatabase): string[] {
  const rows = db.prepare(
    `SELECT id FROM schema_migrations ORDER BY id ASC`
  ).all() as Array<{ id: string }>;
  return rows.map((row) => row.id);
}

export function getRuntimeJournalMode(db: BetterSqliteDatabase): string {
  const row = db.pragma("journal_mode", { simple: true });
  return String(row).toLowerCase();
}
