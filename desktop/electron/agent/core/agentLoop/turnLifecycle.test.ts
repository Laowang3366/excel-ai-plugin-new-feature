import { describe, expect, it, vi } from "vitest";

import type { Thread, TurnItem } from "../../shared/types";
import { ThreadStateManager } from "./threadStateManager";
import { createAgentThread, loadAgentThread } from "./threadLifecycle";
import { completeTurn, createTurn, createUserMessageItem } from "./turnRunner";
import { TurnState } from "./turnState";

describe("turnRunner", () => {
  it("creates and completes a turn", () => {
    const turn = createTurn("thread-1");

    expect(turn).toMatchObject({
      threadId: "thread-1",
      status: "in_progress",
      items: [],
    });
    expect(turn.turnId).toMatch(/^turn-/);

    completeTurn(turn);

    expect(turn.status).toBe("completed");
    expect(turn.completedAt).toEqual(expect.any(Number));
  });

  it("creates a user message item with attachments and client id", () => {
    const item = createUserMessageItem({
      content: "处理这个文件",
      clientId: "client-1",
      attachments: [
        { filePath: "C:/tmp/a.docx", fileName: "a.docx", fileType: "document", size: 123 },
      ],
    });

    expect(item).toMatchObject({
      type: "user_message",
      content: "处理这个文件",
      clientId: "client-1",
      attachments: [
        { filePath: "C:/tmp/a.docx", fileName: "a.docx", fileType: "document", size: 123 },
      ],
    });
  });
});

describe("TurnState", () => {
  it("resets thread-local state while preserving the pending folder id", () => {
    const state = new TurnState();
    state.activeThread = { metadata: { threadId: "t1" }, turns: [] } as unknown as Thread;
    state.activeTurn = createTurn("t1");
    state.compactedHistory = [
      { type: "assistant_message", id: "m1", content: "old", timestamp: 1 },
    ];

    state.resetForNextThread("C:/work");

    expect(state.activeThread).toBeNull();
    expect(state.activeTurn).toBeNull();
    expect(state.compactedHistory).toBeNull();
    expect(state.pendingFolderId).toBe("C:/work");
    expect(state.consumePendingFolderId()).toBe("C:/work");
    expect(state.pendingFolderId).toBeUndefined();
  });
});

describe("ThreadStateManager", () => {
  it("marks an idle loaded thread unloadable after the configured delay", () => {
    const manager = new ThreadStateManager({ idleUnloadMs: 1_000 });

    manager.markLoaded("thread-1", 10_000);

    expect(manager.shouldUnload(10_999)).toBe(false);
    expect(manager.shouldUnload(11_000)).toBe(true);

    manager.markUnloaded(11_000);

    expect(manager.getSnapshot()).toMatchObject({
      status: "unloaded",
      threadId: "thread-1",
      unloadedAt: 11_000,
    });
  });

  it("does not unload a running thread", () => {
    const manager = new ThreadStateManager({ idleUnloadMs: 1_000 });

    manager.markLoaded("thread-1", 10_000);
    manager.markRunning("thread-1", 10_500);

    expect(manager.shouldUnload(20_000)).toBe(false);
  });
});

describe("threadLifecycle", () => {
  it("creates a thread with context window metadata", async () => {
    const thread: Thread = {
      metadata: {
        threadId: "thread-1",
        preview: "",
        modelProvider: "openai",
        model: "model-a",
        createdAt: 1,
        updatedAt: 1,
      },
      turns: [],
    };
    const sessionStore = {
      createThread: vi.fn(async () => thread),
    };

    const result = await createAgentThread({
      sessionStore: sessionStore as any,
      aiConfig: {
        provider: "openai",
        model: "model-a",
        contextWindowSize: 64_000,
        compHash: "chat-v1",
      } as any,
      folderId: "C:/work",
    });

    expect(sessionStore.createThread).toHaveBeenCalledWith("openai", "model-a", "C:/work");
    expect(result.metadata.contextWindowSize).toBe(64_000);
    expect(result.metadata.compHash).toBe("chat-v1");
  });

  it("loads compacted history and re-registers rollout paths", async () => {
    const compactedHistory: TurnItem[] = [
      {
        type: "compacted",
        id: "c1",
        summary: "摘要",
        tokensBefore: 100,
        tokensAfter: 10,
        reason: "auto_token_limit",
        timestamp: 1,
      },
    ];
    const thread: Thread = {
      metadata: {
        threadId: "thread-1",
        preview: "",
        modelProvider: "openai",
        createdAt: 1,
        updatedAt: 1,
        compactedHistory,
      },
      turns: [],
    };
    const sessionStore = {
      loadThread: vi.fn(async () => thread),
      findRolloutPath: vi.fn(async () => "C:/rollout.jsonl"),
      registerRolloutPath: vi.fn(),
    };

    const result = await loadAgentThread(sessionStore as any, "thread-1");

    expect(result).toEqual({ thread, compactedHistory });
    expect(sessionStore.registerRolloutPath).toHaveBeenCalledWith("thread-1", "C:/rollout.jsonl");
  });
});
