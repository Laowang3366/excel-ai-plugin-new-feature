import { describe, expect, it } from "vitest";

import type { Thread, Turn, TurnItem } from "../../shared/types";
import {
  buildContextUsageEvent,
  collectPromptTurnItemGroups,
  collectPromptTurnItems,
} from "./contextUsage";

const userItem = (id: string, content: string): TurnItem => ({
  type: "user_message",
  id,
  content,
  timestamp: 1,
});

function createTurn(turnId: string, items: TurnItem[]): Turn {
  return {
    turnId,
    threadId: "thread-1",
    status: "completed",
    startedAt: 1,
    completedAt: 2,
    items,
  };
}

function createThread(turns: Turn[], contextWindowSize = 1000): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      createdAt: 1,
      updatedAt: 2,
      modelProvider: "test",
      contextWindowSize,
    },
    turns,
  };
}

describe("contextUsage prompt history helpers", () => {
  it("uses compacted history before post-compaction turns and active turn", () => {
    const compacted = [userItem("compact", "summary")];
    const completedTurn = createTurn("turn-1", [userItem("completed", "done")]);
    const activeTurn = createTurn("turn-2", [userItem("active", "now")]);

    const items = collectPromptTurnItems({
      activeThread: createThread([completedTurn]),
      activeTurn,
      compactedHistory: compacted,
    });

    expect(items.map((item) => item.id)).toEqual(["compact", "completed", "active"]);
    expect(
      collectPromptTurnItemGroups({
        activeThread: createThread([completedTurn]),
        activeTurn,
        compactedHistory: compacted,
      }).map((group) => group.map((item) => item.id)),
    ).toEqual([["compact"], ["completed"], ["active"]]);
  });

  it("builds bounded context usage payload from prompt groups", () => {
    const event = buildContextUsageEvent({
      groups: [[userItem("u1", "hello")]],
      activeThread: createThread([], 50),
      compactionConfig: {
        enabled: true,
        contextWindowSize: 50,
        autoCompactTokenThreshold: 40,
        retainedUserMessageMaxTokens: 100,
      },
      systemPrompt: "system",
      tools: [],
    });

    expect(event.type).toBe("context_usage");
    expect(event.contextWindowSize).toBe(50);
    expect(event.threshold).toBe(40);
    expect(event.percentage).toBeGreaterThanOrEqual(0);
    expect(event.percentage).toBeLessThanOrEqual(100);
  });
});
