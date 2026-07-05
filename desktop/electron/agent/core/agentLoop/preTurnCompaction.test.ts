import { describe, expect, it } from "vitest";

import type { Thread, TurnItem } from "../../shared/types";
import { buildPreTurnCompactionPlan } from "./preTurnCompaction";

function createThread(contextWindowSize = 1000): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      modelProvider: "test",
      model: "test-model",
      createdAt: 1,
      updatedAt: 1,
      contextWindowSize,
    },
    turns: [],
  };
}

function createUserItem(content: string): TurnItem {
  return {
    type: "user_message",
    id: "user-1",
    content,
    timestamp: 1,
  };
}

describe("buildPreTurnCompactionPlan", () => {
  it("uses pending compaction reason before automatic threshold checks", () => {
    const plan = buildPreTurnCompactionPlan({
      items: [createUserItem("历史")],
      thread: createThread(),
      globalConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 900,
        retainedUserMessageMaxTokens: 100,
      },
      pendingReason: "model_changed",
    });

    expect(plan.reason).toBe("model_changed");
    expect(plan.config.contextWindowSize).toBe(1000);
  });

  it("ignores pending reason when compaction is disabled", () => {
    const plan = buildPreTurnCompactionPlan({
      items: [createUserItem("历史")],
      thread: createThread(),
      globalConfig: {
        enabled: false,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 1,
        retainedUserMessageMaxTokens: 100,
      },
      pendingReason: "model_changed",
    });

    expect(plan.reason).toBeNull();
  });

  it("returns auto_pre_turn when estimated history exceeds the session threshold", () => {
    const plan = buildPreTurnCompactionPlan({
      items: [createUserItem("很长的历史内容".repeat(200))],
      thread: createThread(2000),
      globalConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 10,
        retainedUserMessageMaxTokens: 100,
      },
      pendingReason: null,
    });

    expect(plan.reason).toBe("auto_pre_turn");
    expect(plan.config.contextWindowSize).toBe(2000);
  });
});
