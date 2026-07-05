import { describe, expect, it, vi } from "vitest";

import type { Thread } from "../../shared/types";
import { ThreadStateManager } from "./threadStateManager";
import { TurnState } from "./turnState";
import {
  resetThreadSession,
  resumeThreadSession,
  startThreadSession,
  sweepIdleThreadSession,
} from "./threadSession";

function createMemorySessionStore() {
  const threads = new Map<string, Thread>();
  return {
    createThread: vi.fn(async (modelProvider: string, model?: string, folderId?: string) => {
      const thread: Thread = {
        metadata: {
          threadId: `thread-${threads.size + 1}`,
          preview: "",
          modelProvider,
          model,
          folderId,
          createdAt: 1,
          updatedAt: 1,
        },
        turns: [],
      };
      threads.set(thread.metadata.threadId, thread);
      return thread;
    }),
    loadThread: vi.fn(async (threadId: string) => threads.get(threadId) ?? null),
    findRolloutPath: vi.fn(async () => null),
    registerRolloutPath: vi.fn(),
    flushRolloutWrites: vi.fn(async () => undefined),
  };
}

describe("threadSession", () => {
  it("resets running sessions after interrupting and stores pending folder", async () => {
    const turnState = new TurnState();
    const manager = new ThreadStateManager();
    const interrupt = vi.fn().mockResolvedValue(undefined);
    const clearIdleUnloadTimer = vi.fn();

    await resetThreadSession({
      isRunning: true,
      interrupt,
      clearIdleUnloadTimer,
      turnState,
      threadStateManager: manager,
      folderId: "folder-1",
    });

    expect(interrupt).toHaveBeenCalled();
    expect(clearIdleUnloadTimer).toHaveBeenCalled();
    expect(turnState.consumePendingFolderId()).toBe("folder-1");
    expect(manager.getSnapshot().status).toBe("not_loaded");
  });

  it("starts and resumes threads while publishing runtime state", async () => {
    const sessionStore = createMemorySessionStore();
    const turnState = new TurnState();
    turnState.pendingFolderId = "folder-1";
    const manager = new ThreadStateManager();
    const activeThreadRef: { current: Thread | null } = { current: null };
    let compactedHistory = undefined as unknown;
    const publishThreadStatus = vi.fn();
    const scheduleIdleThreadUnload = vi.fn();
    const persistThreadSnapshot = vi.fn().mockResolvedValue(undefined);
    const persistThreadRuntime = vi.fn().mockResolvedValue(undefined);

    const threadId = await startThreadSession({
      turnState,
      sessionStore: sessionStore as never,
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
      },
      setActiveThread: (thread) => { activeThreadRef.current = thread; },
      setCompactedHistory: (history) => { compactedHistory = history; },
      threadStateManager: manager,
      publishThreadStatus,
      scheduleIdleThreadUnload,
      persistThreadSnapshot,
      persistThreadRuntime,
    });

    expect(threadId).toBe("thread-1");
    expect(activeThreadRef.current?.metadata.folderId).toBe("folder-1");
    expect(compactedHistory).toBeNull();
    expect(publishThreadStatus).toHaveBeenCalledTimes(1);
    expect(persistThreadRuntime).toHaveBeenCalledWith("thread-1");

    activeThreadRef.current = null;
    const resumed = await resumeThreadSession({
      isRunning: false,
      activeThread: activeThreadRef.current,
      sessionStore: sessionStore as never,
      threadId,
      setActiveThread: (thread) => { activeThreadRef.current = thread; },
      setCompactedHistory: (history) => { compactedHistory = history; },
      threadStateManager: manager,
      publishThreadStatus,
      scheduleIdleThreadUnload,
      persistThreadSnapshot,
      persistThreadRuntime,
    });

    expect(resumed).toBe(true);
    expect((activeThreadRef.current as Thread | null)?.metadata.threadId).toBe("thread-1");
    expect(compactedHistory).toBeNull();
  });

  it("unloads idle active threads after flushing rollout writes", async () => {
    const sessionStore = createMemorySessionStore();
    const manager = new ThreadStateManager({ idleUnloadMs: 10 });
    const thread: Thread = {
      metadata: {
        threadId: "thread-1",
        preview: "",
        modelProvider: "openai",
        createdAt: 1,
        updatedAt: 1,
      },
      turns: [],
    };
    manager.markLoaded("thread-1", 100);
    let activeThread: Thread | null = thread;
    let activeTurn = {} as never;
    let compactedHistory = [{}] as never;
    const publishThreadStatus = vi.fn();
    const clearIdleUnloadTimer = vi.fn();
    const persistThreadRuntime = vi.fn().mockResolvedValue(undefined);

    const unloaded = await sweepIdleThreadSession({
      now: 111,
      isRunning: false,
      activeThread,
      sessionStore: sessionStore as never,
      setActiveThread: (threadValue) => { activeThread = threadValue; },
      setActiveTurn: (turnValue) => { activeTurn = turnValue as never; },
      setCompactedHistory: (history) => { compactedHistory = history as never; },
      threadStateManager: manager,
      publishThreadStatus,
      clearIdleUnloadTimer,
      persistThreadRuntime,
    });

    expect(unloaded).toBe(true);
    expect(sessionStore.flushRolloutWrites).toHaveBeenCalled();
    expect(activeThread).toBeNull();
    expect(activeTurn).toBeNull();
    expect(compactedHistory).toBeNull();
    expect(manager.getSnapshot().status).toBe("unloaded");
    expect(persistThreadRuntime).toHaveBeenCalledWith("thread-1");
  });
});
