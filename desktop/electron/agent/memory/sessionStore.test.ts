/**
 * sessionStore 单元测试
 *
 * 测试 compacted 记录的恢复逻辑：
 * - parseRolloutContent 正确识别 thread 级别的 compacted 历史
 * - 多次压缩时使用最后一次的 replacementHistory
 * - 压缩点之后的 turns 仍被保留
 */

import { mkdir, mkdtemp, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import * as zlib from "zlib";
import { describe, it, expect } from "vitest";
import { SessionStore } from "./sessionStore";
import { StateRuntimeStore } from "./stateRuntimeStore";
import type {
  ThreadId,
  TurnItem,
} from "../shared/types";

/**
 * 通过反射调用 parseRolloutContent（private 方法）
 */
function parseRollout(
  store: SessionStore,
  content: string,
  threadId: ThreadId
): ReturnType<typeof store["parseRolloutContent"]> {
  type RolloutParser = {
    parseRolloutContent: typeof store["parseRolloutContent"];
  };
  return (store as unknown as RolloutParser).parseRolloutContent(content, threadId);
}

/** 构建 JSONL 行 */
function jsonLine(item: Record<string, unknown>): string {
  return JSON.stringify({
    timestamp: new Date().toISOString(),
    item,
  });
}

describe("parseRolloutContent — compacted 记录恢复", () => {
  const store = new SessionStore();
  const threadId = "thread-test-001";

  const userMsg: TurnItem = {
    type: "user_message",
    id: "msg-1",
    content: "Hello",
    timestamp: 1000,
  };

  const assistantMsg: TurnItem = {
    type: "assistant_message",
    id: "asst-1",
    content: "Hi there",
    phase: "final",
    timestamp: 2000,
  };

  const summaryItem: TurnItem = {
    type: "user_message",
    id: "compact-summary-1",
    content: "[compacted] Summary of early conversation",
    timestamp: 3000,
  };

  const postCompactionMsg: TurnItem = {
    type: "user_message",
    id: "msg-2",
    content: "Tell me more",
    timestamp: 4000,
  };

  it("should set metadata.compactedHistory from compacted record", () => {
    const replacementHistory: TurnItem[] = [summaryItem];
    const content = [
      jsonLine({ type: "session_meta", meta: { id: threadId, timestamp: new Date().toISOString(), modelProvider: "test" } }),
      jsonLine({ type: "turn_context", turnId: "turn-1", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: userMsg }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: assistantMsg }),
      jsonLine({ type: "compacted", summary: "Early conv summary", replacementHistory }),
    ].join("\n");

    const thread = parseRollout(store, content, threadId);

    expect(thread.metadata.compactedHistory).toEqual(replacementHistory);
    expect(thread.turns).toHaveLength(0);
  });

  it("should discard pre-compaction turns and keep post-compaction turns", () => {
    const replacementHistory: TurnItem[] = [summaryItem];
    const content = [
      jsonLine({ type: "session_meta", meta: { id: threadId, timestamp: new Date().toISOString(), modelProvider: "test" } }),
      // Turn 1 — should be discarded
      jsonLine({ type: "turn_context", turnId: "turn-1", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: userMsg }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: assistantMsg }),
      // Compaction point
      jsonLine({ type: "compacted", summary: "Early conv summary", replacementHistory }),
      // Turn 2 — should be preserved
      jsonLine({ type: "turn_context", turnId: "turn-2", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-2", item: postCompactionMsg }),
    ].join("\n");

    const thread = parseRollout(store, content, threadId);

    expect(thread.metadata.compactedHistory).toEqual(replacementHistory);
    expect(thread.turns).toHaveLength(1);
    expect(thread.turns[0].turnId).toBe("turn-2");
    expect(thread.turns[0].items).toHaveLength(1);
    expect(thread.turns[0].items[0].id).toBe("msg-2");
  });

  it("should use last compacted record when multiple compactions exist", () => {
    const firstSummary: TurnItem[] = [{ ...summaryItem, content: "[compacted] First" }];
    const secondSummary: TurnItem[] = [{ ...summaryItem, content: "[compacted] Second (latest)" }];
    const content = [
      jsonLine({ type: "session_meta", meta: { id: threadId, timestamp: new Date().toISOString(), modelProvider: "test" } }),
      // Turn 1 — pre-first-compaction
      jsonLine({ type: "turn_context", turnId: "turn-1", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: userMsg }),
      // First compaction
      jsonLine({ type: "compacted", summary: "First", replacementHistory: firstSummary }),
      // Turn 2 — between compactions
      jsonLine({ type: "turn_context", turnId: "turn-2", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-2", item: postCompactionMsg }),
      // Second compaction — should supersede
      jsonLine({ type: "compacted", summary: "Second", replacementHistory: secondSummary }),
    ].join("\n");

    const thread = parseRollout(store, content, threadId);

    expect(thread.metadata.compactedHistory).toEqual(secondSummary);
    // Both turn-1 and turn-2 are discarded by the second compaction
    expect(thread.turns).toHaveLength(0);
  });

  it("should handle session with no compaction (backward compatible)", () => {
    const content = [
      jsonLine({ type: "session_meta", meta: { id: threadId, timestamp: new Date().toISOString(), modelProvider: "test" } }),
      jsonLine({ type: "turn_context", turnId: "turn-1", cwd: "/" }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: userMsg }),
      jsonLine({ type: "turn_item", turnId: "turn-1", item: assistantMsg }),
    ].join("\n");

    const thread = parseRollout(store, content, threadId);

    expect(thread.metadata.compactedHistory).toBeUndefined();
    expect(thread.turns).toHaveLength(1);
    expect(thread.turns[0].items).toHaveLength(2);
  });
});

describe("SessionStore rollout writer", () => {
  it("flushes queued rollout writes before loading a thread", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "session-store-writer-"));
    try {
      const store = new SessionStore(tempDir);
      const thread = await store.createThread("openai", "test-model");
      const item: TurnItem = {
        type: "user_message",
        id: "msg-writer-1",
        content: "异步写入测试",
        timestamp: Date.now(),
      };

      await store.appendTurnItem(thread.metadata.threadId, "turn-writer-1", item);

      const loaded = await store.loadThread(thread.metadata.threadId);

      expect(loaded?.turns).toHaveLength(1);
      expect(loaded?.turns[0].items[0]).toMatchObject({
        type: "user_message",
        content: "异步写入测试",
      });
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("searches compressed rollout archives through the session store facade", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "session-store-compressed-search-"));
    try {
      const sourcePath = path.join(tempDir, "2026", "06", "28", "rollout-old-thread-store.jsonl");
      await mkdir(path.dirname(sourcePath), { recursive: true });
      await writeFile(`${sourcePath}.zst`, await zstdCompress(Buffer.from(`${JSON.stringify({
        timestamp: "2026-06-28T00:00:00.000Z",
        item: {
          type: "turn_item",
          turnId: "turn-store",
          item: {
            type: "assistant_message",
            id: "msg-store",
            phase: "final",
            content: "compressed archive search facade",
            timestamp: 100,
          },
        },
      })}\n`, "utf-8")));

      const store = new SessionStore(tempDir);
      const matches = await store.searchRolloutMatches("archive search", { limit: 5 });

      expect(matches).toEqual([
        expect.objectContaining({
          threadId: "thread-store",
          turnId: "turn-store",
          itemType: "turn_item",
        }),
      ]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("projects rollout writes into StateRuntimeStore logs database", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "session-store-db-"));
    try {
      const stateRuntime = new StateRuntimeStore(path.join(tempDir, "runtime"));
      await stateRuntime.init();

      const store = new SessionStore(path.join(tempDir, "sessions"));
      store.setRolloutEventSink(stateRuntime);

      const thread = await store.createThread("openai", "test-model");
      await store.appendTurnItem(thread.metadata.threadId, "turn-db-1", {
        type: "user_message",
        id: "msg-db-1",
        content: "写入数据库投影",
        timestamp: Date.now(),
      });

      const events = await stateRuntime.listRolloutEvents(thread.metadata.threadId);
      expect(events.map((event) => event.item.type)).toEqual(["session_meta", "turn_item"]);

      await store.flushRolloutWrites();
      await stateRuntime.close();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});

async function zstdCompress(content: Buffer): Promise<Buffer> {
  const compress = (zlib as typeof zlib & {
    zstdCompress?: (buffer: Buffer, callback: (error: Error | null, result: Buffer) => void) => void;
  }).zstdCompress;
  if (!compress) {
    throw new Error("当前 Node 运行时不支持 zstd 压缩");
  }
  return new Promise((resolve, reject) => {
    compress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}
