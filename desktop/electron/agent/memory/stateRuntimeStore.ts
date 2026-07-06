import * as fs from "fs";
import * as path from "path";

import type {
  RolloutItem,
  ThreadId,
  ThreadMetadata,
  ThreadRuntimeSnapshot,
} from "../shared/types";
import {
  mapThreadSnapshot,
} from "./stateRuntimeMappers";
import { extractRolloutSearchContent } from "./rolloutSearchContent";
import {
  defaultStateRuntimeRoot,
  isMemoryRuntime,
  migrateLegacyStateDbIfNeeded,
  resolveRuntimeDatabasePaths,
  runtimeDbNames,
} from "./stateRuntimePaths";
import {
  applyRuntimeMigrations,
  configureRuntimeDatabase,
  getRuntimeJournalMode,
  listAppliedRuntimeMigrations,
  STATE_RUNTIME_MIGRATIONS,
} from "./stateRuntimeSchema";
import { openRuntimeDatabaseWithRecovery } from "./stateRuntimeRecovery";
import { runSqliteTransaction } from "../storage/nodeSqlite";
import {
  getGoalFromDb,
  upsertGoalInDb,
} from "./stateRuntimeGoals";
import {
  archiveLongTermMemoryInDb,
  getLongTermMemoryFromDb,
  getMemoryPipelineCursorFromDb,
  listLongTermMemoriesFromDb,
  listMemoriesFromDb,
  setMemoryPipelineCursorInDb,
  upsertLongTermMemoryInDb,
  upsertMemoryInDb,
} from "./stateRuntimeMemories";
import {
  appendRolloutItemsToLogs,
  listRolloutEventsFromLogs,
  searchRolloutMatchesInLogs,
} from "./stateRuntimeRolloutEvents";
import {
  appendToolExecutionLogToLogs,
  listToolExecutionLogsFromLogs,
} from "./stateRuntimeToolLogs";
import type {
  RuntimeConnections,
  RuntimeDatabasePaths,
  RuntimeDbName,
  RuntimeGoalRecord,
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryListOptions,
  RuntimeMemoryRecord,
  RuntimeRecoveryReport,
  RuntimeRolloutEvent,
  RuntimeRolloutSearchMatch,
  RuntimeToolExecutionLogRecord,
} from "./stateRuntimeTypes";

export type {
  RuntimeDatabasePaths,
  RuntimeDbName,
  RuntimeGoalRecord,
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryListOptions,
  RuntimeMemoryRecord,
  RuntimeRolloutEvent,
  RuntimeRolloutSearchMatch,
  RuntimeToolExecutionLogRecord,
} from "./stateRuntimeTypes";

/**
 * SQLite 状态运行时。
 *
 * 关联模块：
 * - stateRuntimePaths.ts: 解析四库路径，并迁移旧的 state-runtime.db。
 * - stateRuntimeSchema.ts: 维护四库 schema、迁移和 WAL 配置。
 * - sessionStore.ts: 仍写 JSONL 兼容审计副本，同时把 rollout 事件投影到 logs.db。
 * - core/agentLoop: 写入线程快照和运行态，避免仅靠 JSONL 回放恢复状态。
 */
export class StateRuntimeStore {
  private dbs: RuntimeConnections | null = null;
  private readonly dbPaths: RuntimeDatabasePaths;
  private readonly legacyStateDbPath?: string;
  private readonly recoveryReports: RuntimeRecoveryReport[] = [];
  private initialized = false;
  private transactionDepth = 0;

  constructor(runtimeRoot = defaultStateRuntimeRoot()) {
    const resolved = resolveRuntimeDatabasePaths(runtimeRoot);
    this.dbPaths = resolved.dbPaths;
    this.legacyStateDbPath = resolved.legacyStateDbPath;
  }

  async init(): Promise<void> {
    if (this.initialized) return;

    if (!isMemoryRuntime(this.dbPaths)) {
      await fs.promises.mkdir(path.dirname(this.dbPaths.state), { recursive: true });
      await migrateLegacyStateDbIfNeeded(this.dbPaths.state, this.legacyStateDbPath);
    }

    this.dbs = {
      state: openRuntimeDatabaseWithRecovery(this.dbPaths.state, "state", this.recoveryReports),
      logs: openRuntimeDatabaseWithRecovery(this.dbPaths.logs, "logs", this.recoveryReports),
      goals: openRuntimeDatabaseWithRecovery(this.dbPaths.goals, "goals", this.recoveryReports),
      memories: openRuntimeDatabaseWithRecovery(this.dbPaths.memories, "memories", this.recoveryReports),
    };

    for (const name of runtimeDbNames()) {
      configureRuntimeDatabase(this.dbs[name]);
      applyRuntimeMigrations(this.dbs[name], STATE_RUNTIME_MIGRATIONS[name]);
    }
    this.initialized = true;
    this.backfillDerivedIndexes();
  }

  async close(): Promise<void> {
    if (!this.initialized || !this.dbs) return;
    for (const name of runtimeDbNames()) {
      this.dbs[name].close();
    }
    this.dbs = null;
    this.initialized = false;
    this.transactionDepth = 0;
  }

  getDatabasePaths(): RuntimeDatabasePaths {
    return { ...this.dbPaths };
  }

  getRecoveryReports(): RuntimeRecoveryReport[] {
    return this.recoveryReports.map((report) => ({
      ...report,
      backupPaths: [...report.backupPaths],
    }));
  }

  getJournalModes(): Record<RuntimeDbName, string> {
    const dbs = this.getDbs();
    return {
      state: getRuntimeJournalMode(dbs.state),
      logs: getRuntimeJournalMode(dbs.logs),
      goals: getRuntimeJournalMode(dbs.goals),
      memories: getRuntimeJournalMode(dbs.memories),
    };
  }

  getAppliedMigrations(): Record<RuntimeDbName, string[]> {
    const dbs = this.getDbs();
    return {
      state: listAppliedRuntimeMigrations(dbs.state),
      logs: listAppliedRuntimeMigrations(dbs.logs),
      goals: listAppliedRuntimeMigrations(dbs.goals),
      memories: listAppliedRuntimeMigrations(dbs.memories),
    };
  }

  async upsertThreadSnapshot(metadata: ThreadMetadata): Promise<void> {
    const db = this.getDbs().state;
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
    if (metadata.name) {
      await this.appendThreadName(metadata.threadId, metadata.name, metadata.updatedAt);
    }
  }

  async getThreadSnapshot(threadId: ThreadId): Promise<ThreadMetadata | null> {
    const row = this.getDbs().state.prepare(
      `SELECT * FROM thread_snapshots WHERE thread_id = ?`
    ).get(threadId) as Record<string, any> | undefined;
    return row ? mapThreadSnapshot(row) : null;
  }

  async listThreadSnapshots(): Promise<ThreadMetadata[]> {
    const rows = this.getDbs().state.prepare(
      `SELECT * FROM thread_snapshots ORDER BY updated_at DESC`
    ).all() as Record<string, any>[];
    return rows.map(mapThreadSnapshot);
  }

  async updateThreadRuntime(snapshot: ThreadRuntimeSnapshot & { threadId: ThreadId }): Promise<void> {
    this.getDbs().state.prepare(
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

  async getThreadRuntime(threadId: ThreadId): Promise<(ThreadRuntimeSnapshot & { threadId: ThreadId }) | null> {
    const row = this.getDbs().state.prepare(
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

  async appendRolloutItems(threadId: ThreadId, items: RolloutItem[]): Promise<void> {
    appendRolloutItemsToLogs(this.getDbs().logs, threadId, items, (write) => this.runLogsWrite(write));
  }

  async listRolloutEvents(threadId: ThreadId): Promise<RuntimeRolloutEvent[]> {
    return listRolloutEventsFromLogs(this.getDbs().logs, threadId);
  }

  async searchRolloutMatches(
    query: string,
    options: { limit?: number } = {}
  ): Promise<RuntimeRolloutSearchMatch[]> {
    return searchRolloutMatchesInLogs(this.getDbs().logs, query, options);
  }

  async appendToolExecutionLog(record: RuntimeToolExecutionLogRecord): Promise<void> {
    appendToolExecutionLogToLogs(this.getDbs().logs, record);
  }

  async listToolExecutionLogs(
    threadId: ThreadId,
    options: { limit?: number } = {}
  ): Promise<RuntimeToolExecutionLogRecord[]> {
    return listToolExecutionLogsFromLogs(this.getDbs().logs, threadId, options);
  }

  async appendThreadName(threadId: ThreadId, name: string, updatedAt = Date.now()): Promise<void> {
    const trimmed = name.trim();
    if (!trimmed) return;

    this.getDbs().state.prepare(
      `INSERT INTO thread_names (thread_id, name, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(thread_id) DO UPDATE SET
         name = excluded.name,
         updated_at = excluded.updated_at`
    ).run(threadId, trimmed, updatedAt);
  }

  async findThreadNameByIdStr(threadId: string): Promise<string | null> {
    const row = this.getDbs().state.prepare(
      `SELECT name FROM thread_names WHERE thread_id = ?`
    ).get(threadId) as { name: string } | undefined;
    return row?.name ?? null;
  }

  async upsertGoal(goal: RuntimeGoalRecord): Promise<void> {
    upsertGoalInDb(this.getDbs().goals, goal);
  }

  async getGoal(goalId: string): Promise<RuntimeGoalRecord | null> {
    return getGoalFromDb(this.getDbs().goals, goalId);
  }

  async upsertMemory(memory: RuntimeMemoryRecord): Promise<void> {
    upsertMemoryInDb(this.getDbs().memories, memory);
  }

  async listMemories(namespace?: string): Promise<RuntimeMemoryRecord[]> {
    return listMemoriesFromDb(this.getDbs().memories, namespace);
  }

  async upsertLongTermMemory(memory: RuntimeLongTermMemoryRecord): Promise<void> {
    upsertLongTermMemoryInDb(this.getDbs().memories, memory);
  }

  async getLongTermMemory(memoryId: string): Promise<RuntimeLongTermMemoryRecord | null> {
    return getLongTermMemoryFromDb(this.getDbs().memories, memoryId);
  }

  async archiveLongTermMemory(memoryId: string, updatedAt = Date.now()): Promise<RuntimeLongTermMemoryRecord | null> {
    return archiveLongTermMemoryInDb(this.getDbs().memories, memoryId, updatedAt);
  }

  async listLongTermMemories(
    options: RuntimeMemoryListOptions = {}
  ): Promise<RuntimeLongTermMemoryRecord[]> {
    return listLongTermMemoriesFromDb(this.getDbs().memories, options);
  }

  async getMemoryPipelineCursor(pipelineId: string): Promise<number> {
    return getMemoryPipelineCursorFromDb(this.getDbs().memories, pipelineId);
  }

  async setMemoryPipelineCursor(pipelineId: string, lastEventId: number): Promise<void> {
    setMemoryPipelineCursorInDb(this.getDbs().memories, pipelineId, lastEventId);
  }

  async transaction<T>(fn: (tx: StateRuntimeStore) => Promise<T>): Promise<T> {
    if (this.transactionDepth > 0) {
      return fn(this);
    }

    const dbs = this.getDbs();
    const begun: RuntimeDbName[] = [];
    for (const name of runtimeDbNames()) {
      dbs[name].prepare("BEGIN").run();
      begun.push(name);
    }
    this.transactionDepth = 1;
    try {
      const result = await fn(this);
      for (const name of runtimeDbNames()) {
        dbs[name].prepare("COMMIT").run();
      }
      return result;
    } catch (error) {
      for (const name of [...begun].reverse()) {
        try {
          dbs[name].prepare("ROLLBACK").run();
        } catch {
          // SQLite may already have closed the transaction after a failed COMMIT.
        }
      }
      throw error;
    } finally {
      this.transactionDepth = 0;
    }
  }

  private getDbs(): RuntimeConnections {
    if (!this.initialized || !this.dbs) {
      throw new Error("StateRuntimeStore 尚未初始化");
    }
    return this.dbs;
  }

  private runLogsWrite<T>(fn: () => T): T {
    if (this.transactionDepth > 0) return fn();
    return runSqliteTransaction(this.getDbs().logs, fn);
  }

  private backfillDerivedIndexes(): void {
    const dbs = this.getDbs();
    dbs.state.prepare(
      `INSERT OR IGNORE INTO thread_names (thread_id, name, updated_at)
       SELECT thread_id, name, updated_at
       FROM thread_snapshots
       WHERE name IS NOT NULL AND trim(name) <> ''`
    ).run();

    const indexedIds = new Set(
      (dbs.logs.prepare(`SELECT rowid AS id FROM rollout_events_fts`).all() as Array<{ id: number }>)
        .map((row) => row.id)
    );
    const missingRows = dbs.logs.prepare(
      `SELECT id, thread_id, turn_id, item_type, item_json
       FROM rollout_events
       ORDER BY id ASC`
    ).all() as Record<string, any>[];
    const insertSearch = dbs.logs.prepare(
      `INSERT OR IGNORE INTO rollout_events_fts (rowid, thread_id, turn_id, item_type, content, item_json)
       VALUES (?, ?, ?, ?, ?, ?)`
    );
    const write = () => this.runLogsWrite(() => {
      for (const row of missingRows) {
        if (indexedIds.has(row.id)) continue;
        let content = row.item_json;
        try {
          content = extractRolloutSearchContent(JSON.parse(row.item_json));
        } catch {
          // 损坏 JSONL 投影仍保留原始 JSON 供粗略检索。
        }
        insertSearch.run(row.id, row.thread_id, row.turn_id ?? null, row.item_type, content, row.item_json);
      }
    });
    write();
  }
}
