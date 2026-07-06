import * as fs from "fs";
import * as path from "path";

import type {
  RolloutItem,
  RolloutLine,
  ThreadId,
  ThreadMetadata,
  ThreadRuntimeSnapshot,
} from "../shared/types";
import {
  buildRolloutFtsQuery,
  clampMemoryListOffset,
  getRolloutTurnId,
  mapGoal,
  mapLongTermMemory,
  mapMemory,
  mapThreadSnapshot,
  mapToolExecutionLog,
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
import { clampNumber } from "../shared/numberLimits";
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
    if (items.length === 0) return;
    const insert = this.getDbs().logs.prepare(
      `INSERT INTO rollout_events (thread_id, turn_id, item_type, timestamp, item_json)
       VALUES (?, ?, ?, ?, ?)`
    );
    const insertSearch = this.getDbs().logs.prepare(
      `INSERT INTO rollout_events_fts (rowid, thread_id, turn_id, item_type, content, item_json)
       VALUES (?, ?, ?, ?, ?, ?)`
    );
    const write = (rows: RolloutLine[]) => this.runLogsWrite(() => {
      for (const line of rows) {
        const itemJson = JSON.stringify(line.item);
        const result = insert.run(
          threadId,
          getRolloutTurnId(line.item) ?? null,
          line.item.type,
          line.timestamp,
          itemJson
        );
        insertSearch.run(
          Number(result.lastInsertRowid),
          threadId,
          getRolloutTurnId(line.item) ?? null,
          line.item.type,
          extractRolloutSearchContent(line.item),
          itemJson
        );
      }
    });

    write(items.map((item) => ({ timestamp: new Date().toISOString(), item })));
  }

  async listRolloutEvents(threadId: ThreadId): Promise<RuntimeRolloutEvent[]> {
    const rows = this.getDbs().logs.prepare(
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

  async searchRolloutMatches(
    query: string,
    options: { limit?: number } = {}
  ): Promise<RuntimeRolloutSearchMatch[]> {
    const normalized = query.trim();
    if (!normalized) return [];

    const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
    const rows = this.getDbs().logs.prepare(
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

  async appendToolExecutionLog(record: RuntimeToolExecutionLogRecord): Promise<void> {
    this.getDbs().logs.prepare(
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
      record.metadata ? JSON.stringify(record.metadata) : null
    );
  }

  async listToolExecutionLogs(
    threadId: ThreadId,
    options: { limit?: number } = {}
  ): Promise<RuntimeToolExecutionLogRecord[]> {
    const rows = this.getDbs().logs.prepare(
      `SELECT * FROM tool_execution_logs
       WHERE thread_id = ?
       ORDER BY id ASC
       LIMIT ?`
    ).all(threadId, clampNumber(options.limit, { fallback: 200, min: 1, max: 1000 })) as Record<string, any>[];
    return rows.map(mapToolExecutionLog);
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
    this.getDbs().goals.prepare(
      `INSERT INTO goals (
        goal_id, thread_id, objective, status, token_budget, token_usage,
        created_at, updated_at, completed_at, payload_json
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(goal_id) DO UPDATE SET
        thread_id = excluded.thread_id,
        objective = excluded.objective,
        status = excluded.status,
        token_budget = excluded.token_budget,
        token_usage = excluded.token_usage,
        updated_at = excluded.updated_at,
        completed_at = excluded.completed_at,
        payload_json = excluded.payload_json`
    ).run(
      goal.goalId,
      goal.threadId ?? null,
      goal.objective,
      goal.status,
      goal.tokenBudget ?? null,
      goal.tokenUsage ?? null,
      goal.createdAt,
      goal.updatedAt,
      goal.completedAt ?? null,
      goal.payload ? JSON.stringify(goal.payload) : null
    );
  }

  async getGoal(goalId: string): Promise<RuntimeGoalRecord | null> {
    const row = this.getDbs().goals.prepare(
      `SELECT * FROM goals WHERE goal_id = ?`
    ).get(goalId) as Record<string, any> | undefined;
    return row ? mapGoal(row) : null;
  }

  async upsertMemory(memory: RuntimeMemoryRecord): Promise<void> {
    this.getDbs().memories.prepare(
      `INSERT INTO memories (
        memory_id, namespace, content, metadata_json, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?)
      ON CONFLICT(memory_id) DO UPDATE SET
        namespace = excluded.namespace,
        content = excluded.content,
        metadata_json = excluded.metadata_json,
        updated_at = excluded.updated_at`
    ).run(
      memory.memoryId,
      memory.namespace,
      memory.content,
      memory.metadata ? JSON.stringify(memory.metadata) : null,
      memory.createdAt,
      memory.updatedAt
    );
  }

  async listMemories(namespace?: string): Promise<RuntimeMemoryRecord[]> {
    const rows = namespace
      ? this.getDbs().memories.prepare(
          `SELECT * FROM memories WHERE namespace = ? ORDER BY updated_at DESC`
        ).all(namespace)
      : this.getDbs().memories.prepare(
          `SELECT * FROM memories ORDER BY updated_at DESC`
        ).all();
    return (rows as Record<string, any>[]).map(mapMemory);
  }

  async upsertLongTermMemory(memory: RuntimeLongTermMemoryRecord): Promise<void> {
    this.getDbs().memories.prepare(
      `INSERT INTO long_term_memories (
        memory_id, namespace, kind, visibility, status, content, summary, confidence,
        source_thread_id, source_event_id, workspace_fingerprint, expires_at,
        metadata_json, citations_json, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(memory_id) DO UPDATE SET
        namespace = excluded.namespace,
        kind = excluded.kind,
        visibility = excluded.visibility,
        status = excluded.status,
        content = excluded.content,
        summary = excluded.summary,
        confidence = excluded.confidence,
        source_thread_id = excluded.source_thread_id,
        source_event_id = excluded.source_event_id,
        workspace_fingerprint = excluded.workspace_fingerprint,
        expires_at = excluded.expires_at,
        metadata_json = excluded.metadata_json,
        citations_json = excluded.citations_json,
        updated_at = excluded.updated_at`
    ).run(
      memory.memoryId,
      memory.namespace,
      memory.kind,
      memory.visibility,
      memory.status,
      memory.content,
      memory.summary ?? null,
      memory.confidence ?? null,
      memory.sourceThreadId ?? null,
      memory.sourceEventId ?? null,
      memory.workspaceFingerprint ?? null,
      memory.expiresAt ?? null,
      memory.metadata ? JSON.stringify(memory.metadata) : null,
      memory.citations ? JSON.stringify(memory.citations) : null,
      memory.createdAt,
      memory.updatedAt
    );
  }

  async getLongTermMemory(memoryId: string): Promise<RuntimeLongTermMemoryRecord | null> {
    const row = this.getDbs().memories.prepare(
      `SELECT * FROM long_term_memories WHERE memory_id = ?`
    ).get(memoryId) as Record<string, any> | undefined;
    return row ? mapLongTermMemory(row) : null;
  }

  async archiveLongTermMemory(memoryId: string, updatedAt = Date.now()): Promise<RuntimeLongTermMemoryRecord | null> {
    const existing = await this.getLongTermMemory(memoryId);
    if (!existing) return null;

    this.getDbs().memories.prepare(
      `UPDATE long_term_memories
       SET status = 'archived', updated_at = ?
       WHERE memory_id = ?`
    ).run(updatedAt, memoryId);

    return this.getLongTermMemory(memoryId);
  }

  async listLongTermMemories(
    options: RuntimeMemoryListOptions = {}
  ): Promise<RuntimeLongTermMemoryRecord[]> {
    const filters: string[] = [];
    const params: Array<string | number> = [];

    if (options.namespace) {
      filters.push("namespace = ?");
      params.push(options.namespace);
    }
    if (options.kind) {
      filters.push("kind = ?");
      params.push(options.kind);
    }
    if (options.visibility) {
      filters.push("visibility = ?");
      params.push(options.visibility);
    }
    if (options.status) {
      filters.push("status = ?");
      params.push(options.status);
    }

    const where = filters.length > 0 ? `WHERE ${filters.join(" AND ")}` : "";
    const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
    const offset = clampMemoryListOffset(options.offset);
    const rows = this.getDbs().memories.prepare(
      `SELECT * FROM long_term_memories ${where} ORDER BY updated_at DESC LIMIT ? OFFSET ?`
    ).all(...params, limit, offset) as Record<string, any>[];
    return rows.map(mapLongTermMemory);
  }

  async getMemoryPipelineCursor(pipelineId: string): Promise<number> {
    const row = this.getDbs().memories.prepare(
      `SELECT last_event_id FROM memory_pipeline_state WHERE pipeline_id = ?`
    ).get(pipelineId) as { last_event_id: number } | undefined;
    return row?.last_event_id ?? 0;
  }

  async setMemoryPipelineCursor(pipelineId: string, lastEventId: number): Promise<void> {
    this.getDbs().memories.prepare(
      `INSERT INTO memory_pipeline_state (pipeline_id, last_event_id, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(pipeline_id) DO UPDATE SET
         last_event_id = excluded.last_event_id,
         updated_at = excluded.updated_at`
    ).run(pipelineId, lastEventId, Date.now());
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
