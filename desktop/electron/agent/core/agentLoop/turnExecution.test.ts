import { describe, expect, it, vi } from "vitest";

import type { AgentTurnCallbacks, Thread, Turn } from "../../shared/types";
import { ThreadStateManager } from "./threadStateManager";
import { TurnState } from "./turnState";
import {
  beginTurnRun,
  completeSuccessfulTurn,
  createStartedTurn,
  finishTurnRun,
  handleTurnFailure,
  prepareThreadForTurn,
} from "./turnExecution";

function createThread(): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      modelProvider: "openai",
      createdAt: 1,
      updatedAt: 1,
    },
    turns: [],
  };
}

function createCallbacks(events: unknown[]): AgentTurnCallbacks {
  return {
    onEvent: (event) => events.push(event),
  };
}

describe("turnExecution", () => {
  it("begins a run and prepares an active thread", async () => {
    const turnState = new TurnState();
    const thread = createThread();
    const manager = new ThreadStateManager();
    const publishThreadStatus = vi.fn();
    const persistThreadRuntime = vi.fn().mockResolvedValue(undefined);
    const boundCallbacks = createCallbacks([]);

    beginTurnRun(turnState);
    turnState.activeThread = thread;
    const result = await prepareThreadForTurn({
      turnState,
      startThread: vi.fn(),
      clearIdleUnloadTimer: vi.fn(),
      threadStateManager: manager,
      publishThreadStatus,
      persistThreadRuntime,
      bindCallbacksToThread: vi.fn(() => boundCallbacks),
      callbacks: createCallbacks([]),
      clientId: "client-1",
    });

    expect(turnState.isRunning).toBe(true);
    expect(turnState.abortController).toBeInstanceOf(AbortController);
    expect(result.thread).toBe(thread);
    expect(result.callbacks).toBe(boundCallbacks);
    expect(manager.getSnapshot()).toMatchObject({ status: "running", threadId: "thread-1" });
    expect(persistThreadRuntime).toHaveBeenCalledWith("thread-1");
    expect(publishThreadStatus).toHaveBeenCalled();
  });

  it("creates and completes a started turn with user message events", async () => {
    const turnState = new TurnState();
    const thread = createThread();
    const events: unknown[] = [];
    const sessionStore = {
      appendTurnItem: vi.fn().mockResolvedValue(undefined),
      appendTurnUsage: vi.fn().mockResolvedValue(undefined),
    };
    const persistThreadSnapshot = vi.fn().mockResolvedValue(undefined);
    const scheduleTurnMemoryExtraction = vi.fn();

    const turn = await createStartedTurn({
      turnInput: { content: "你好" },
      thread,
      turnState,
      callbacks: createCallbacks(events),
      sessionStore: sessionStore as never,
      persistThreadSnapshot,
    });
    turn.tokenUsage = { inputTokens: 1, outputTokens: 2 };

    await completeSuccessfulTurn({
      thread,
      turn,
      callbacks: createCallbacks(events),
      sessionStore: sessionStore as never,
      persistThreadSnapshot,
      scheduleTurnMemoryExtraction,
    });

    expect(turnState.activeTurn).toBe(turn);
    expect(thread.metadata.preview).toBe("你好");
    expect(thread.metadata.activeTurnId).toBeUndefined();
    expect(thread.turns).toEqual([turn]);
    expect(sessionStore.appendTurnItem).toHaveBeenCalled();
    expect(sessionStore.appendTurnUsage).toHaveBeenCalledWith("thread-1", turn.turnId, turn.tokenUsage);
    expect(scheduleTurnMemoryExtraction).toHaveBeenCalledWith(thread, turn);
    expect(events.map((event) => (event as { type: string }).type)).toContain("turn_completed");
  });

  it("records failed turns and finalizes running state", async () => {
    const turnState = new TurnState();
    const thread = createThread();
    const turn: Turn = {
      threadId: "thread-1",
      turnId: "turn-1",
      status: "in_progress",
      startedAt: 1,
      items: [],
    };
    const events: unknown[] = [];
    const manager = new ThreadStateManager();
    turnState.activeThread = thread;
    turnState.activeTurn = turn;
    turnState.isRunning = true;

    await handleTurnFailure({
      error: new Error("boom"),
      turnState,
      callbacks: createCallbacks(events),
      persistThreadSnapshot: vi.fn().mockResolvedValue(undefined),
    });
    await finishTurnRun({
      turnState,
      threadStateManager: manager,
      publishThreadStatus: vi.fn(),
      scheduleIdleThreadUnload: vi.fn(),
      persistThreadRuntime: vi.fn().mockResolvedValue(undefined),
    });

    expect(turn.status).toBe("failed");
    expect(thread.turns).toEqual([turn]);
    expect(turnState.isRunning).toBe(false);
    expect(turnState.abortController).toBeNull();
    expect(events[0]).toMatchObject({ type: "turn_failed", error: "boom" });
  });
});
