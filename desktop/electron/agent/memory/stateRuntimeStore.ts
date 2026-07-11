import * as fs from "fs";
import * as path from "path";

import type { RolloutItem, ThreadId, ThreadMetadata, ThreadRuntimeSnapshot } from "../shared/types";
import { extractRolloutSearchContent } from "./rolloutSearchContent";
import {
  defaultStateRuntimeRoot,
  isMemoryRuntime,
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
import { getGoalFromDb, upsertGoalInDb } from "./stateRuntimeGoals";
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
import {
  getThreadRuntimeFromDb,
  getThreadSnapshotFromDb,
  deleteThreadStateFromDb,
  listThreadSnapshotsFromDb,
  updateThreadRuntimeInDb,
  upsertThreadSnapshotInDb,
} from "./stateRuntimeThreads";
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
 * - stateRuntimePaths.ts: 解析四库路径。
 * - stateRuntimeSchema.ts: 维护四库 schema、迁移和 WAL 配置。
 * - sessionStore.ts: 仍写 JSONL 兼容审计副本，同时把 rollout 事件投影到 logs.db。
 * - core/agentLoop: 写入线程快照和运行态，避免仅靠 JSONL 回放恢复状态。
 */
export class StateRuntimeStore {
  private dbs: RuntimeConnections | null = null;
  private readonly dbPaths: RuntimeDatabasePaths;
  private readonly recoveryReports: RuntimeRecoveryReport[] = [];
  private initialized = false;

  constructor(runtimeRoot = defaultStateRuntimeRoot()) {
    const resolved = resolveRuntimeDatabasePaths(runtimeRoot);
    this.dbPaths = resolved.dbPaths;
  }

  async init(): Promise<void> {
    if (this.initialized) return;

    if (!isMemoryRuntime(this.dbPaths)) {
      await fs.promises.mkdir(path.dirname(this.dbPaths.state), { recursive: true });
    }

    this.dbs = {
      state: openRuntimeDatabaseWithRecovery(this.dbPaths.state, "state", this.recoveryReports),
      logs: openRuntimeDatabaseWithRecovery(this.dbPaths.logs, "logs", this.recoveryReports),
      goals: openRuntimeDatabaseWithRecovery(this.dbPaths.goals, "goals", this.recoveryReports),
      memories: openRuntimeDatabaseWithRecovery(
        this.dbPaths.memories,
        "memories",
        this.recoveryReports,
      ),
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
    upsertThreadSnapshotInDb(this.getDbs().state, metadata);
    if (metadata.name) {
      await this.appendThreadName(metadata.threadId, metadata.name, metadata.updatedAt);
    } else {
      this.getDbs()
        .state.prepare("DELETE FROM thread_names WHERE thread_id = ?")
        .run(metadata.threadId);
    }
  }

  async getThreadSnapshot(threadId: ThreadId): Promise<ThreadMetadata | null> {
    return getThreadSnapshotFromDb(this.getDbs().state, threadId);
  }

  async listThreadSnapshots(): Promise<ThreadMetadata[]> {
    return listThreadSnapshotsFromDb(this.getDbs().state);
  }

  async deleteThreadData(threadId: ThreadId): Promise<void> {
    deleteThreadStateFromDb(this.getDbs().state, threadId);
    runSqliteTransaction(this.getDbs().logs, () => {
      this.getDbs()
        .logs.prepare("DELETE FROM rollout_events_fts WHERE thread_id = ?")
        .run(threadId);
      this.getDbs().logs.prepare("DELETE FROM rollout_events WHERE thread_id = ?").run(threadId);
      this.getDbs()
        .logs.prepare("DELETE FROM tool_execution_logs WHERE thread_id = ?")
        .run(threadId);
    });
  }

  async updateThreadRuntime(
    snapshot: ThreadRuntimeSnapshot & { threadId: ThreadId },
  ): Promise<void> {
    updateThreadRuntimeInDb(this.getDbs().state, snapshot);
  }

  async getThreadRuntime(
    threadId: ThreadId,
  ): Promise<(ThreadRuntimeSnapshot & { threadId: ThreadId }) | null> {
    return getThreadRuntimeFromDb(this.getDbs().state, threadId);
  }

  async appendRolloutItems(threadId: ThreadId, items: RolloutItem[]): Promise<void> {
    appendRolloutItemsToLogs(this.getDbs().logs, threadId, items, (write) =>
      this.runLogsWrite(write),
    );
  }

  async listRolloutEvents(threadId: ThreadId): Promise<RuntimeRolloutEvent[]> {
    return listRolloutEventsFromLogs(this.getDbs().logs, threadId);
  }

  async searchRolloutMatches(
    query: string,
    options: { limit?: number } = {},
  ): Promise<RuntimeRolloutSearchMatch[]> {
    return searchRolloutMatchesInLogs(this.getDbs().logs, query, options);
  }

  async appendToolExecutionLog(record: RuntimeToolExecutionLogRecord): Promise<void> {
    appendToolExecutionLogToLogs(this.getDbs().logs, record);
  }

  async listToolExecutionLogs(
    threadId: ThreadId,
    options: { limit?: number } = {},
  ): Promise<RuntimeToolExecutionLogRecord[]> {
    return listToolExecutionLogsFromLogs(this.getDbs().logs, threadId, options);
  }

  async appendThreadName(threadId: ThreadId, name: string, updatedAt = Date.now()): Promise<void> {
    const trimmed = name.trim();
    if (!trimmed) return;

    this.getDbs()
      .state.prepare(
        `INSERT INTO thread_names (thread_id, name, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(thread_id) DO UPDATE SET
         name = excluded.name,
         updated_at = excluded.updated_at`,
      )
      .run(threadId, trimmed, updatedAt);
  }

  async findThreadNameByIdStr(threadId: string): Promise<string | null> {
    const row = this.getDbs()
      .state.prepare(`SELECT name FROM thread_names WHERE thread_id = ?`)
      .get(threadId) as { name: string } | undefined;
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

  async archiveLongTermMemory(
    memoryId: string,
    updatedAt = Date.now(),
  ): Promise<RuntimeLongTermMemoryRecord | null> {
    return archiveLongTermMemoryInDb(this.getDbs().memories, memoryId, updatedAt);
  }

  async listLongTermMemories(
    options: RuntimeMemoryListOptions = {},
  ): Promise<RuntimeLongTermMemoryRecord[]> {
    return listLongTermMemoriesFromDb(this.getDbs().memories, options);
  }

  async getMemoryPipelineCursor(pipelineId: string): Promise<number> {
    return getMemoryPipelineCursorFromDb(this.getDbs().memories, pipelineId);
  }

  async setMemoryPipelineCursor(pipelineId: string, lastEventId: number): Promise<void> {
    setMemoryPipelineCursorInDb(this.getDbs().memories, pipelineId, lastEventId);
  }

  private getDbs(): RuntimeConnections {
    if (!this.initialized || !this.dbs) {
      throw new Error("StateRuntimeStore 尚未初始化");
    }
    return this.dbs;
  }

  private runLogsWrite<T>(fn: () => T): T {
    return runSqliteTransaction(this.getDbs().logs, fn);
  }

  private backfillDerivedIndexes(): void {
    const dbs = this.getDbs();
    dbs.state
      .prepare(
        `INSERT OR IGNORE INTO thread_names (thread_id, name, updated_at)
       SELECT thread_id, name, updated_at
       FROM thread_snapshots
       WHERE name IS NOT NULL AND trim(name) <> ''`,
      )
      .run();

    const indexedIds = new Set(
      (
        dbs.logs.prepare(`SELECT rowid AS id FROM rollout_events_fts`).all() as Array<{
          id: number;
        }>
      ).map((row) => row.id),
    );
    const missingRows = dbs.logs
      .prepare(
        `SELECT id, thread_id, turn_id, item_type, item_json
       FROM rollout_events
       ORDER BY id ASC`,
      )
      .all() as Record<string, any>[];
    const insertSearch = dbs.logs.prepare(
      `INSERT OR IGNORE INTO rollout_events_fts (rowid, thread_id, turn_id, item_type, content, item_json)
       VALUES (?, ?, ?, ?, ?, ?)`,
    );
    const write = () =>
      this.runLogsWrite(() => {
        for (const row of missingRows) {
          if (indexedIds.has(row.id)) continue;
          let content = row.item_json;
          try {
            content = extractRolloutSearchContent(JSON.parse(row.item_json));
          } catch {
            // 损坏 JSONL 投影仍保留原始 JSON 供粗略检索。
          }
          insertSearch.run(
            row.id,
            row.thread_id,
            row.turn_id ?? null,
            row.item_type,
            content,
            row.item_json,
          );
        }
      });
    write();
  }
}
