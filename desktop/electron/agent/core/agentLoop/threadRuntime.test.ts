import { describe, expect, it, vi } from "vitest";

import type { AgentTurnCallbacks, Thread, Turn } from "../../shared/types";
import {
  attachRolloutEventSink,
  bindCallbacksToThread,
  persistThreadRuntime,
  persistThreadSnapshot,
  scheduleTurnMemoryExtraction,
} from "./threadRuntime";

function createThread(): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      createdAt: 1,
      updatedAt: 1,
      modelProvider: "test",
    },
    turns: [],
  };
}

function createTurn(status: Turn["status"]): Turn {
  return {
    turnId: "turn-1",
    threadId: "thread-1",
    status,
    startedAt: 1,
    items: [],
  };
}

describe("threadRuntime helpers", () => {
  it("binds thread and client ids to events and stream deltas", () => {
    const events: unknown[] = [];
    const streamDeltas: unknown[][] = [];
    const callbacks: AgentTurnCallbacks = {
      onEvent: (event) => events.push(event),
      onStreamDelta: (...args) => streamDeltas.push(args),
    };

    const bound = bindCallbacksToThread({
      callbacks,
      threadId: "thread-1",
      clientId: "client-1",
    });
    bound.onEvent({ type: "turn_started", turnId: "turn-1" });
    bound.onStreamDelta?.("delta", "assistant_message", 2);

    expect(events[0]).toMatchObject({
      type: "turn_started",
      threadId: "thread-1",
      clientId: "client-1",
    });
    expect(streamDeltas[0]).toEqual(["delta", "assistant_message", 2, "thread-1", "client-1"]);
  });

  it("persists thread snapshot and runtime when stores are present", async () => {
    const stateRuntimeStore = {
      upsertThreadSnapshot: vi.fn().mockResolvedValue(undefined),
      updateThreadRuntime: vi.fn().mockResolvedValue(undefined),
    };

    await persistThreadSnapshot({
      stateRuntimeStore: stateRuntimeStore as never,
      thread: createThread(),
    });
    await persistThreadRuntime({
      stateRuntimeStore: stateRuntimeStore as never,
      snapshot: {
        status: "active",
        threadId: "thread-1",
        lastActiveAt: 10,
        idleUnloadMs: 100,
      },
      threadId: "thread-1",
    });

    expect(stateRuntimeStore.upsertThreadSnapshot).toHaveBeenCalledWith(
      expect.objectContaining({ threadId: "thread-1" }),
    );
    expect(stateRuntimeStore.updateThreadRuntime).toHaveBeenCalledWith(
      expect.objectContaining({ threadId: "thread-1", status: "active" }),
    );
  });

  it("attaches rollout sink and only schedules memory extraction for completed turns", () => {
    const sessionStore = {
      setRolloutEventSink: vi.fn(),
    };
    const memoryStore = {
      updateRuntime: vi.fn(),
    };

    attachRolloutEventSink({
      sessionStore: sessionStore as never,
      stateRuntimeStore: memoryStore as never,
    });
    scheduleTurnMemoryExtraction({
      aiClient: {} as never,
      memoryStore: undefined,
      thread: createThread(),
      turn: createTurn("failed"),
    });

    expect(sessionStore.setRolloutEventSink).toHaveBeenCalledWith(memoryStore);
  });
});
