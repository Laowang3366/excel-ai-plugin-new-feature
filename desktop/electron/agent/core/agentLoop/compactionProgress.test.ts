import { describe, expect, it, vi } from "vitest";

import type { AgentTurnCallbacks, TurnItem } from "../../shared/types";
import type { SessionStore } from "../../memory/sessionStore";
import {
  archiveRolloutIfConfigured,
  completeCompactionProgress,
  failCompactionProgress,
  startCompactionProgress,
} from "./compactionProgress";

function createCallbacks(events: unknown[]): AgentTurnCallbacks {
  return {
    onEvent: (event) => events.push(event),
  };
}

function createSessionStore() {
  return {
    appendRolloutItems: vi.fn().mockResolvedValue(undefined),
    spawnRolloutCompressionWorker: vi.fn().mockResolvedValue(undefined),
  } as unknown as SessionStore & {
    appendRolloutItems: ReturnType<typeof vi.fn>;
    spawnRolloutCompressionWorker: ReturnType<typeof vi.fn>;
  };
}

const items: TurnItem[] = [
  {
    type: "user_message",
    id: "user-1",
    content: "请总结",
    timestamp: 1,
  },
];

describe("compactionProgress", () => {
  it("emits start and completion events while recording rollout params", async () => {
    const sessionStore = createSessionStore();
    const events: unknown[] = [];

    const progress = await startCompactionProgress({
      sessionStore,
      threadId: "thread-1",
      reason: "auto_pre_turn",
      items,
      callbacks: createCallbacks(events),
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
        summaryRetryCount: 2,
      },
    });

    completeCompactionProgress({
      progress,
      tokensBefore: 120,
      tokensAfter: 40,
      summary: "摘要",
      callbacks: createCallbacks(events),
    });

    expect(sessionStore.appendRolloutItems).toHaveBeenCalledWith("thread-1", [
      expect.objectContaining({
        type: "compact_params",
        reason: "auto_pre_turn",
        status: "started",
        itemCount: 1,
      }),
    ]);
    expect(events[0]).toMatchObject({
      type: "thread_compact_started",
      threadId: "thread-1",
      params: {
        reason: "auto_pre_turn",
        tokenThreshold: 800,
        contextWindowSize: 1000,
        retryCount: 2,
      },
    });
    expect(events[1]).toMatchObject({
      type: "item_started",
      item: {
        type: "compact_progress",
        status: "running",
      },
    });
    expect(events[2]).toMatchObject({
      type: "item_completed",
      item: {
        type: "compact_progress",
        status: "completed",
        tokensBefore: 120,
        tokensAfter: 40,
        summary: "摘要",
      },
    });
  });

  it("records failed progress and keeps archive optional", async () => {
    const sessionStore = createSessionStore();
    const events: unknown[] = [];
    const progress = await startCompactionProgress({
      sessionStore,
      threadId: "thread-1",
      reason: "auto_token_limit",
      items,
      callbacks: createCallbacks([]),
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
      },
    });

    await failCompactionProgress({
      sessionStore,
      threadId: "thread-1",
      progress,
      items,
      error: new Error("boom"),
      callbacks: createCallbacks(events),
    });
    await archiveRolloutIfConfigured({
      sessionStore,
      threadId: "thread-1",
      threshold: 0,
    });

    expect(sessionStore.appendRolloutItems).toHaveBeenLastCalledWith("thread-1", [
      expect.objectContaining({
        type: "compact_params",
        reason: "auto_token_limit",
        status: "failed",
        error: "boom",
      }),
    ]);
    expect(events[0]).toMatchObject({
      type: "item_completed",
      item: {
        type: "compact_progress",
        status: "failed",
        message: "上下文压缩失败：boom",
      },
    });
    expect(sessionStore.spawnRolloutCompressionWorker).not.toHaveBeenCalled();
  });
});
