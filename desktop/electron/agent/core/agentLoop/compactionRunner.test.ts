import { describe, expect, it, vi } from "vitest";

import type {
  AgentTurnCallbacks,
  CompactProgressItem,
  Thread,
  Turn,
  TurnItem,
} from "../../shared/types";
import type { SessionStore } from "../../memory/sessionStore";
import { runAutoCompaction, runMidTurnCompaction } from "./compactionRunner";

const userItem = (id: string, content: string): TurnItem => ({
  type: "user_message",
  id,
  content,
  timestamp: 1,
});

const assistantItem = (id: string, content: string): TurnItem => ({
  type: "assistant_message",
  id,
  content,
  phase: "commentary",
  timestamp: 1,
});

function createThread(items: TurnItem[] = []): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      createdAt: 1,
      updatedAt: 1,
      modelProvider: "test",
    },
    turns:
      items.length > 0
        ? [
            {
              threadId: "thread-1",
              turnId: "turn-old",
              status: "completed",
              startedAt: 1,
              completedAt: 2,
              items,
            },
          ]
        : [],
  };
}

function createProgress(reason: CompactProgressItem["reason"]): CompactProgressItem {
  return {
    type: "compact_progress",
    id: "progress-1",
    reason,
    status: "running",
    message: "running",
    tokensBefore: 10,
    timestamp: 1,
  };
}

function createDeps(input: {
  allItems: TurnItem[];
  activeThread: Thread | null;
  compactedHistory: TurnItem[][];
  events: unknown[];
}) {
  const sessionStore = {
    appendRolloutItems: vi.fn().mockResolvedValue(undefined),
  } as unknown as SessionStore & { appendRolloutItems: ReturnType<typeof vi.fn> };
  return {
    sessionStore,
    getAllTurnItems: () => input.allItems,
    generateCompactionSummary: vi.fn().mockResolvedValue("压缩摘要"),
    startCompactionProgress: vi
      .fn()
      .mockImplementation((_threadId, reason) => Promise.resolve(createProgress(reason))),
    completeCompactionProgress: vi
      .fn()
      .mockImplementation((_progress, tokensBefore, tokensAfter, summary, callbacks) => {
        callbacks.onEvent({
          type: "item_completed",
          item: {
            type: "compact_progress",
            status: "completed",
            tokensBefore,
            tokensAfter,
            summary,
          },
        });
      }),
    failCompactionProgress: vi.fn().mockResolvedValue(undefined),
    archiveRolloutIfConfigured: vi.fn().mockResolvedValue(undefined),
    setCompactedHistory: (history: TurnItem[]) => input.compactedHistory.push(history),
    getActiveThread: () => input.activeThread,
    compactionConfig: {
      enabled: true,
      contextWindowSize: 1000,
      autoCompactTokenThreshold: 800,
      retainedUserMessageMaxTokens: 100,
    },
  };
}

function createCallbacks(events: unknown[]): AgentTurnCallbacks {
  return {
    onEvent: (event) => events.push(event),
  };
}

describe("compactionRunner", () => {
  it("runs pre-turn compaction and clears completed turns", async () => {
    const allItems = [userItem("u1", "历史消息")];
    const thread = createThread(allItems);
    const events: unknown[] = [];
    const compactedHistory: TurnItem[][] = [];
    const deps = createDeps({ allItems, activeThread: thread, compactedHistory, events });

    await runAutoCompaction({
      thread,
      reason: "auto_pre_turn",
      callbacks: createCallbacks(events),
      deps,
    });

    expect(thread.turns).toEqual([]);
    expect(compactedHistory).toHaveLength(1);
    expect(deps.sessionStore.appendRolloutItems).toHaveBeenCalledWith("thread-1", [
      expect.objectContaining({ type: "compacted", summary: "压缩摘要" }),
      expect.objectContaining({
        type: "compact_params",
        reason: "auto_pre_turn",
        status: "completed",
      }),
    ]);
    expect(events).toEqual([
      expect.objectContaining({
        type: "item_completed",
        item: expect.objectContaining({ type: "compact_progress", status: "completed" }),
      }),
      expect.objectContaining({ type: "context_compacted", summary: "压缩摘要" }),
    ]);
  });

  it("runs mid-turn compaction while keeping current user messages visible", async () => {
    const currentUser = userItem("u-current", "当前问题");
    const intermediate = assistantItem("a1", "中间过程");
    const turn: Turn = {
      threadId: "thread-1",
      turnId: "turn-1",
      status: "in_progress",
      startedAt: 1,
      items: [currentUser, intermediate],
    };
    const allItems = [userItem("u-old", "历史"), currentUser, intermediate];
    const thread = createThread([userItem("u-old", "历史")]);
    const events: unknown[] = [];
    const compactedHistory: TurnItem[][] = [];
    const deps = createDeps({ allItems, activeThread: thread, compactedHistory, events });

    await runMidTurnCompaction({
      turn,
      callbacks: createCallbacks(events),
      deps,
    });

    expect(turn.items).toEqual([currentUser]);
    expect(thread.turns).toEqual([]);
    expect(compactedHistory[0].some((item) => item.id === "u-current")).toBe(false);
    expect(deps.archiveRolloutIfConfigured).toHaveBeenCalledWith("thread-1");
    expect(events.some((event) => (event as { type?: string }).type === "context_compacted")).toBe(
      true,
    );
  });

  it("keeps pre-turn history unchanged when summary generation fails", async () => {
    const allItems = [userItem("u1", "历史消息")];
    const thread = createThread(allItems);
    const originalTurns = thread.turns;
    const events: unknown[] = [];
    const compactedHistory: TurnItem[][] = [];
    const deps = createDeps({ allItems, activeThread: thread, compactedHistory, events });
    const failure = new Error("summary failed");
    deps.generateCompactionSummary.mockRejectedValueOnce(failure);

    await expect(
      runAutoCompaction({
        thread,
        reason: "auto_pre_turn",
        callbacks: createCallbacks(events),
        deps,
      }),
    ).rejects.toBe(failure);

    expect(thread.turns).toBe(originalTurns);
    expect(compactedHistory).toEqual([]);
    expect(deps.sessionStore.appendRolloutItems).not.toHaveBeenCalled();
    expect(deps.archiveRolloutIfConfigured).not.toHaveBeenCalled();
    expect(deps.completeCompactionProgress).not.toHaveBeenCalled();
    expect(deps.failCompactionProgress).toHaveBeenCalledOnce();
  });

  it("keeps current turn items unchanged when mid-turn summary generation fails", async () => {
    const currentUser = userItem("u-current", "当前问题");
    const intermediate = assistantItem("a1", "中间过程");
    const turn: Turn = {
      threadId: "thread-1",
      turnId: "turn-1",
      status: "in_progress",
      startedAt: 1,
      items: [currentUser, intermediate],
    };
    const originalItems = turn.items;
    const allItems = [userItem("u-old", "历史"), currentUser, intermediate];
    const thread = createThread([userItem("u-old", "历史")]);
    const originalTurns = thread.turns;
    const events: unknown[] = [];
    const compactedHistory: TurnItem[][] = [];
    const deps = createDeps({ allItems, activeThread: thread, compactedHistory, events });
    const failure = new Error("summary failed");
    deps.generateCompactionSummary.mockRejectedValueOnce(failure);

    await expect(
      runMidTurnCompaction({
        turn,
        callbacks: createCallbacks(events),
        deps,
      }),
    ).rejects.toBe(failure);

    expect(turn.items).toBe(originalItems);
    expect(thread.turns).toBe(originalTurns);
    expect(compactedHistory).toEqual([]);
    expect(deps.sessionStore.appendRolloutItems).not.toHaveBeenCalled();
    expect(deps.archiveRolloutIfConfigured).not.toHaveBeenCalled();
    expect(deps.completeCompactionProgress).not.toHaveBeenCalled();
    expect(deps.failCompactionProgress).toHaveBeenCalledOnce();
  });
});
