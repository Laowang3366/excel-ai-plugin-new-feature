import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it } from "vitest";

import type { ThreadMetadata } from "../shared/types";
import { openSqliteDatabase } from "../storage/nodeSqlite";
import { StateRuntimeStore } from "./stateRuntimeStore";

function createMetadata(patch: Partial<ThreadMetadata> = {}): ThreadMetadata {
  return {
    threadId: "thread-1",
    preview: "预览",
    modelProvider: "openai",
    model: "model-a",
    createdAt: 1,
    updatedAt: 2,
    folderId: "folder-a",
    ...patch,
  };
}

describe("StateRuntimeStore", () => {
  it("initializes four sqlite databases with migrations and WAL", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();

      expect(store.getDatabasePaths()).toMatchObject({
        state: path.join(tempDir, "state.db"),
        logs: path.join(tempDir, "logs.db"),
        goals: path.join(tempDir, "goals.db"),
        memories: path.join(tempDir, "memories.db"),
      });
      expect(store.getJournalModes()).toEqual({
        state: "wal",
        logs: "wal",
        goals: "wal",
        memories: "wal",
      });
      expect(store.getAppliedMigrations().state).toContain("001_state_runtime");
      expect(store.getAppliedMigrations().logs).toContain("001_rollout_logs");
      expect(store.getAppliedMigrations().goals).toContain("001_goals");
      expect(store.getAppliedMigrations().memories).toContain("001_memories");

      await store.close();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("backs up and replaces a corrupted runtime database during init", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    try {
      const stateDbPath = path.join(tempDir, "state.db");
      const originalContent = Buffer.from("not a sqlite database");
      await writeFile(stateDbPath, originalContent);
      await writeFile(`${stateDbPath}-wal`, "broken wal");

      store = new StateRuntimeStore(tempDir);
      await store.init();

      const reports = store.getRecoveryReports();
      expect(reports).toHaveLength(1);
      expect(reports[0]).toMatchObject({
        dbName: "state",
        dbPath: stateDbPath,
      });
      expect(reports[0].backupPaths.length).toBeGreaterThanOrEqual(1);
      await expect(readFile(reports[0].backupPaths[0])).resolves.toEqual(originalContent);
      expect(store.getAppliedMigrations().state).toContain("001_state_runtime");
    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("persists thread metadata snapshots and runtime status", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    try {
      const store = new StateRuntimeStore(tempDir);
      await store.init();

      await store.upsertThreadSnapshot(createMetadata());
      await store.updateThreadRuntime({
        threadId: "thread-1",
        status: "running",
        lastActiveAt: 10,
        idleUnloadMs: 30_000,
      });
      await store.close();

      const reloaded = new StateRuntimeStore(tempDir);
      await reloaded.init();

      expect(await reloaded.getThreadSnapshot("thread-1")).toMatchObject({
        threadId: "thread-1",
        preview: "预览",
        modelProvider: "openai",
        model: "model-a",
        folderId: "folder-a",
      });
      expect(await reloaded.getThreadRuntime("thread-1")).toMatchObject({
        threadId: "thread-1",
        status: "running",
        lastActiveAt: 10,
        idleUnloadMs: 30_000,
      });
      await reloaded.close();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("lists thread snapshots from sqlite in updated order", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    await store.upsertThreadSnapshot(createMetadata({
      threadId: "thread-old",
      preview: "旧会话",
      updatedAt: 10,
    }));
    await store.upsertThreadSnapshot(createMetadata({
      threadId: "thread-new",
      preview: "新会话",
      updatedAt: 20,
    }));

    expect((await store.listThreadSnapshots()).map((thread) => thread.threadId)).toEqual([
      "thread-new",
      "thread-old",
    ]);
    await store.close();
  });

  it("stores rollout events in logs database as primary query data", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    try {
      const store = new StateRuntimeStore(tempDir);
      await store.init();

      await store.appendRolloutItems("thread-logs", [
        {
          type: "turn_item",
          turnId: "turn-1",
          item: {
            type: "user_message",
            id: "msg-1",
            content: "数据库日志",
            timestamp: 100,
          },
        },
        {
          type: "turn_usage",
          turnId: "turn-1",
          usage: { inputTokens: 3, outputTokens: 5 },
        },
      ]);

      const events = await store.listRolloutEvents("thread-logs");
      expect(events.map((event) => event.item.type)).toEqual(["turn_item", "turn_usage"]);
      expect(events[0]).toMatchObject({
        threadId: "thread-logs",
        turnId: "turn-1",
        itemType: "turn_item",
      });

      await store.close();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("stores structured tool execution logs in the logs database", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    await store.appendToolExecutionLog({
      threadId: "thread-tools",
      turnId: "turn-tools",
      toolCallId: "call-tools",
      toolName: "office.action.apply",
      status: "success",
      durationMs: 12,
      timestamp: 1234,
      argumentsSummary: "{\"app\":\"presentation\"}",
      resultSummary: "{\"status\":\"success\"}",
      error: undefined,
    });

    expect(await store.listToolExecutionLogs("thread-tools")).toEqual([
      expect.objectContaining({
        threadId: "thread-tools",
        turnId: "turn-tools",
        toolCallId: "call-tools",
        toolName: "office.action.apply",
        status: "success",
        durationMs: 12,
        argumentsSummary: "{\"app\":\"presentation\"}",
        resultSummary: "{\"status\":\"success\"}",
      }),
    ]);
    await store.close();
  });

  it("indexes thread names and finds them by thread id string", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();

      await store.appendThreadName("thread-name-1", "健康饮食宣传 PPT", 100);

      expect(await store.findThreadNameByIdStr("thread-name-1")).toBe("健康饮食宣传 PPT");
      expect(await store.findThreadNameByIdStr("missing-thread")).toBeNull();

    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("searches rollout matches across threads from the logs database", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();

      await store.appendRolloutItems("thread-search-1", [
        {
          type: "turn_item",
          turnId: "turn-a",
          item: {
            type: "user_message",
            id: "msg-a",
            content: "Need a wellness presentation for students",
            timestamp: 100,
          },
        },
      ]);
      await store.appendRolloutItems("thread-search-2", [
        {
          type: "turn_item",
          turnId: "turn-b",
          item: {
            type: "assistant_message",
            id: "msg-b",
            phase: "final",
            content: "Inventory report is ready",
            timestamp: 200,
          },
        },
      ]);

      const matches = await store.searchRolloutMatches("wellness presentation", { limit: 5 });

      expect(matches).toEqual([
        expect.objectContaining({
          threadId: "thread-search-1",
          turnId: "turn-a",
          itemType: "turn_item",
        }),
      ]);
      expect(matches[0].snippet).toContain("wellness");

    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("keeps secrets and raw tool payloads out of the rollout FTS projection", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    const canary = "sk-1234567890abcdefghijklmnop";
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();
      await store.appendRolloutItems("thread-redaction", [
        {
          type: "turn_item",
          turnId: "turn-redaction",
          item: {
            type: "user_message",
            id: "msg-redaction",
            content: `quarterly planning ${canary}`,
            timestamp: 100,
          },
        },
        {
          type: "turn_item",
          turnId: "turn-redaction",
          item: {
            type: "tool_call",
            id: "call-redaction",
            toolName: "web.search",
            arguments: { query: `customer forecast ${canary}` },
            status: "completed",
            timestamp: 101,
          },
        },
        {
          type: "turn_item",
          turnId: "turn-redaction",
          item: {
            type: "reasoning",
            id: "reasoning-redaction",
            summaryText: ["safe planning summary"],
            rawContent: [`private chain ${canary}`],
            timestamp: 102,
          },
        },
      ]);
      const dbPath = store.getDatabasePaths().logs;
      await store.close();
      store = undefined;

      const logsDb = openSqliteDatabase(dbPath);
      const ftsRows = logsDb.prepare(
        `SELECT content FROM rollout_events_fts ORDER BY rowid`,
      ).all() as Array<{ content: string }>;
      const ftsSchema = logsDb.prepare(`PRAGMA table_info(rollout_events_fts)`).all() as Array<{
        name: string;
      }>;
      logsDb.close();

      const serializedFts = JSON.stringify(ftsRows);
      expect(serializedFts).not.toContain(canary);
      expect(serializedFts).not.toContain("customer forecast");
      expect(serializedFts).not.toContain("private chain");
      expect(serializedFts).toContain("quarterly planning");
      expect(serializedFts).toContain("safe planning summary");
      expect(ftsSchema.map((column) => column.name)).not.toContain("item_json");
    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("rebuilds legacy rollout FTS rows without their duplicated raw item JSON", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    const canary = "sk-1234567890abcdefghijklmnop";
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();
      await store.appendRolloutItems("thread-legacy-fts", [
        {
          type: "turn_item",
          turnId: "turn-legacy-fts",
          item: {
            type: "user_message",
            id: "msg-legacy-fts",
            content: `legacy planning ${canary}`,
            timestamp: 100,
          },
        },
      ]);
      const logsPath = store.getDatabasePaths().logs;
      await store.close();
      store = undefined;

      const legacyDb = openSqliteDatabase(logsPath);
      const event = legacyDb.prepare(
        `SELECT id, thread_id, turn_id, item_type, item_json FROM rollout_events LIMIT 1`,
      ).get() as Record<string, any>;
      legacyDb.exec(`
        DROP TABLE rollout_events_fts;
        CREATE VIRTUAL TABLE rollout_events_fts USING fts5(
          thread_id UNINDEXED,
          turn_id UNINDEXED,
          item_type UNINDEXED,
          content,
          item_json UNINDEXED
        );
        DELETE FROM schema_migrations WHERE id = '006_minimize_rollout_fts';
      `);
      legacyDb.prepare(
        `INSERT INTO rollout_events_fts
          (rowid, thread_id, turn_id, item_type, content, item_json)
         VALUES (?, ?, ?, ?, ?, ?)`,
      ).run(
        event.id,
        event.thread_id,
        event.turn_id,
        event.item_type,
        `legacy planning ${canary}`,
        event.item_json,
      );
      legacyDb.close();

      store = new StateRuntimeStore(tempDir);
      await store.init();
      await store.close();
      store = undefined;

      const migratedDb = openSqliteDatabase(logsPath);
      const content = migratedDb.prepare(
        `SELECT content FROM rollout_events_fts LIMIT 1`,
      ).get() as { content: string };
      const columns = migratedDb.prepare(`PRAGMA table_info(rollout_events_fts)`).all() as Array<{
        name: string;
      }>;
      migratedDb.close();

      expect(content.content).toContain("legacy planning");
      expect(content.content).not.toContain(canary);
      expect(columns.map((column) => column.name)).not.toContain("item_json");
    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("backfills derived thread name and rollout search indexes on init", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    let store: StateRuntimeStore | undefined;
    try {
      store = new StateRuntimeStore(tempDir);
      await store.init();
      await store.upsertThreadSnapshot(createMetadata({
        threadId: "thread-backfill",
        name: "回填测试会话",
      }));
      await store.appendRolloutItems("thread-backfill", [
        {
          type: "turn_item",
          turnId: "turn-backfill",
          item: {
            type: "user_message",
            id: "msg-backfill",
            content: "legacy searchable text",
            timestamp: 300,
          },
        },
      ]);
      const dbPaths = store.getDatabasePaths();
      await store.close();

      const stateDb = openSqliteDatabase(dbPaths.state);
      stateDb.prepare(`DELETE FROM thread_names`).run();
      stateDb.close();
      const logsDb = openSqliteDatabase(dbPaths.logs);
      logsDb.prepare(`DELETE FROM rollout_events_fts`).run();
      logsDb.close();

      store = new StateRuntimeStore(tempDir);
      await store.init();

      expect(await store.findThreadNameByIdStr("thread-backfill")).toBe("回填测试会话");
      expect(await store.searchRolloutMatches("legacy searchable", { limit: 10 })).toEqual([
        expect.objectContaining({ threadId: "thread-backfill" }),
      ]);
    } finally {
      await store?.close();
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("stores goals and memories in dedicated databases", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "state-runtime-"));
    try {
      const store = new StateRuntimeStore(tempDir);
      await store.init();

      await store.upsertGoal({
        goalId: "goal-1",
        threadId: "thread-1",
        objective: "完成数据库迁移",
        status: "active",
        createdAt: 10,
        updatedAt: 20,
      });
      await store.upsertMemory({
        memoryId: "memory-1",
        namespace: "project",
        content: "优先 Open XML，COM 兜底",
        metadata: { source: "user" },
        createdAt: 30,
        updatedAt: 40,
      });

      expect(await store.getGoal("goal-1")).toMatchObject({
        goalId: "goal-1",
        objective: "完成数据库迁移",
        status: "active",
      });
      expect(await store.listMemories("project")).toEqual([
        expect.objectContaining({
          memoryId: "memory-1",
          content: "优先 Open XML，COM 兜底",
          metadata: { source: "user" },
        }),
      ]);

      await store.close();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("persists structured long-term memories with visibility fields", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    await store.upsertLongTermMemory({
      memoryId: "mem-1",
      namespace: "global",
      kind: "preference",
      visibility: "user",
      status: "active",
      content: "回复先给结论",
      summary: "回复风格偏好",
      confidence: 0.9,
      citations: [{ threadId: "thread-1", eventId: 1 }],
      metadata: { source: "tool" },
      createdAt: 100,
      updatedAt: 100,
    });

    expect(await store.listLongTermMemories({ visibility: "user" })).toMatchObject([
      {
        memoryId: "mem-1",
        kind: "preference",
        visibility: "user",
        status: "active",
        content: "回复先给结论",
      },
    ]);
    expect(await store.listLongTermMemories({ visibility: "internal" })).toEqual([]);
    await store.close();
  });

  it("applies offset when listing long-term memories", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    await store.upsertLongTermMemory({
      memoryId: "mem-newer",
      namespace: "global",
      kind: "preference",
      visibility: "user",
      status: "active",
      content: "newer memory",
      createdAt: 200,
      updatedAt: 200,
    });
    await store.upsertLongTermMemory({
      memoryId: "mem-older",
      namespace: "global",
      kind: "preference",
      visibility: "user",
      status: "active",
      content: "older memory",
      createdAt: 100,
      updatedAt: 100,
    });

    expect(
      (await store.listLongTermMemories({ limit: 1, offset: 1 })).map(
        (memory) => memory.memoryId,
      ),
    ).toEqual(["mem-older"]);
    expect(
      (await store.listLongTermMemories({ limit: 1, offset: -1 })).map(
        (memory) => memory.memoryId,
      ),
    ).toEqual(["mem-newer"]);
    expect(
      (await store.listLongTermMemories({
        limit: 1,
        offset: Number.NaN,
      })).map((memory) => memory.memoryId),
    ).toEqual(["mem-newer"]);
    await store.close();
  });

  it("archives long-term memories by id", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    await store.upsertLongTermMemory({
      memoryId: "mem-delete",
      namespace: "global",
      kind: "preference",
      visibility: "user",
      status: "active",
      content: "delete me",
      createdAt: 100,
      updatedAt: 100,
    });

    const archived = await store.archiveLongTermMemory("mem-delete", 200);

    expect(archived).toMatchObject({
      memoryId: "mem-delete",
      status: "archived",
      updatedAt: 200,
    });
    expect(await store.listLongTermMemories({ status: "active" })).toEqual([]);
    expect((await store.listLongTermMemories({ status: "archived" }))[0].memoryId).toBe("mem-delete");
    expect(await store.archiveLongTermMemory("missing")).toBeNull();
    await store.close();
  });

  it("tracks memory pipeline cursor in memories database", async () => {
    const store = new StateRuntimeStore(":memory:");
    await store.init();

    expect(await store.getMemoryPipelineCursor("default")).toBe(0);
    await store.setMemoryPipelineCursor("default", 42);
    expect(await store.getMemoryPipelineCursor("default")).toBe(42);
    await store.close();
  });
});
