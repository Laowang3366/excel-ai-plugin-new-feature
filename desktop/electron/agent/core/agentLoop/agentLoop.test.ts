/**
 * agentLoop/streamCollector 单元测试
 */

import { describe, it, expect, vi, beforeEach } from "vitest";

const aiClientMocks = vi.hoisted(() => ({
  client: {
    streamChat: vi.fn(),
    chat: vi.fn(),
  },
  streamChatParams: [] as any[],
}));

vi.mock("../../providers/aiClient", async () => {
  const actual = await vi.importActual<typeof import("../../providers/aiClient")>(
    "../../providers/aiClient"
  );
  return {
    ...actual,
    createAIClient: vi.fn(() => aiClientMocks.client),
  };
});

import { AgentLoop } from "./agentLoop";
import { ThreadWatchManager } from "./threadWatchManager";
import {
  shouldRequireApproval,
  clearAlwaysAllowedTools,
  markToolAlwaysAllowed,
  getAlwaysAllowedTools,
  executeTool,
  processToolCalls,
} from "./toolExecutor";
import type {
  AgentTurnCallbacks,
  ToolCallItem,
  ToolExecutor,
  Turn,
  TurnItem,
  Thread,
} from "../../shared/types";
import type { AIStreamEvent, StreamChatParams } from "../../providers/aiClient";
import type { ToolCallInfo } from "./streamCollector";

async function* streamEvents(events: AIStreamEvent[]): AsyncIterable<AIStreamEvent> {
  for (const event of events) {
    yield event;
  }
}

function createMemorySessionStore(): any {
  const threads = new Map<string, Thread>();
  return {
    createThread: vi.fn(async (modelProvider: string, model?: string, folderId?: string) => {
      const threadId = `thread-${threads.size + 1}`;
      const now = Date.now();
      const thread: Thread = {
        metadata: {
          threadId,
          preview: "",
          modelProvider,
          model,
          createdAt: now,
          updatedAt: now,
          folderId,
        },
        turns: [],
      };
      threads.set(threadId, thread);
      return thread;
    }),
    appendRolloutItems: vi.fn(async () => {}),
    appendTurnItem: vi.fn(async () => {}),
    appendTurnUsage: vi.fn(async () => {}),
    flushRolloutWrites: vi.fn(async () => {}),
    spawnRolloutCompressionWorker: vi.fn(async () => ({ compressed: [], skipped: [] })),
    loadThread: vi.fn(async (threadId: string) => threads.get(threadId) ?? null),
    findRolloutPath: vi.fn(async () => null),
    registerRolloutPath: vi.fn(),
  };
}

// ============================================================
// executeTool
// ============================================================

describe("executeTool", () => {
  it("should convert async executor failures into tool results", async () => {
    const executors = new Map<string, ToolExecutor>([
      [
        "workbook.inspect",
        {
          name: "workbook.inspect",
          execute: vi.fn(async () => {
            throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");
          }),
        },
      ],
    ]);

    const result = await executeTool("workbook.inspect", "{}", executors);

    expect(result).toEqual({
      success: false,
      error: "工具执行错误: 未连接到 Excel/WPS，请先在侧边栏点击连接",
    });
  });
});

// ============================================================
// processToolCalls
// ============================================================

describe("processToolCalls", () => {
  it("should mark a pending tool call as failed when async execution fails", async () => {
    const activeItem: ToolCallItem = {
      type: "tool_call",
      id: "call-1",
      toolName: "workbook.inspect",
      arguments: {},
      status: "pending",
      timestamp: 1000,
    };
    const turn: Turn = {
      turnId: "turn-1",
      threadId: "thread-1",
      status: "in_progress",
      items: [activeItem],
      startedAt: 1000,
    };
    const toolCalls: ToolCallInfo[] = [
      { id: "call-1", name: "workbook.inspect", arguments: "{}" },
    ];
    const executors = new Map<string, ToolExecutor>([
      [
        "workbook.inspect",
        {
          name: "workbook.inspect",
          execute: vi.fn(async () => {
            throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");
          }),
        },
      ],
    ]);
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const appended: TurnItem[] = [];
    const sessionStoreAppend = vi.fn(
      async (_threadId: string, _turnId: string, item: TurnItem) => {
        appended.push(item);
      }
    );

    await processToolCalls(
      toolCalls,
      new Map([["call-1", activeItem]]),
      turn,
      executors,
      { permissionMode: "confirm_all" },
      callbacks,
      sessionStoreAppend
    );

    const resultItem = turn.items.find(
      (item) => item.type === "tool_result" && item.toolCallId === "call-1"
    );

    expect(activeItem.status).toBe("failed");
    expect(resultItem).toMatchObject({
      type: "tool_result",
      toolCallId: "call-1",
      toolName: "workbook.inspect",
      result: "工具执行错误: 未连接到 Excel/WPS，请先在侧边栏点击连接",
      isError: true,
    });
    expect(appended).toContain(resultItem);
  });

  it("confirm_all: follows user permission for non-forced shell prompts", async () => {
    const activeItem: ToolCallItem = {
      type: "tool_call",
      id: "call-shell",
      toolName: "shell.execute",
      arguments: { command: "curl https://example.com" },
      status: "pending",
      timestamp: 1000,
    };
    const turn: Turn = {
      turnId: "turn-1",
      threadId: "thread-1",
      status: "in_progress",
      items: [activeItem],
      startedAt: 1000,
    };
    const executor = vi.fn(async () => ({ success: true, data: { stdout: "ok" } }));
    const executors = new Map<string, ToolExecutor>([
      ["shell.execute", { name: "shell.execute", execute: executor }],
    ]);
    const requestToolApproval = vi.fn(async () => ({ approved: true }));
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const sessionStoreAppend = vi.fn(async () => {});

    await processToolCalls(
      [{ id: "call-shell", name: "shell.execute", arguments: "{\"command\":\"curl https://example.com\"}" }],
      new Map([["call-shell", activeItem]]),
      turn,
      executors,
      { permissionMode: "confirm_all", requestToolApproval },
      callbacks,
      sessionStoreAppend
    );

    expect(requestToolApproval).not.toHaveBeenCalled();
    expect(executor).toHaveBeenCalledWith(
      expect.objectContaining({ command: "curl https://example.com" }),
      expect.objectContaining({
        sandboxEvaluation: expect.objectContaining({ decision: "allow" }),
      })
    );
  });

  it("confirm_all: still asks for forced shell prompts", async () => {
    const activeItem: ToolCallItem = {
      type: "tool_call",
      id: "call-nohup",
      toolName: "shell.execute",
      arguments: { command: "nohup node server.js" },
      status: "pending",
      timestamp: 1000,
    };
    const turn: Turn = {
      turnId: "turn-1",
      threadId: "thread-1",
      status: "in_progress",
      items: [activeItem],
      startedAt: 1000,
    };
    const executor = vi.fn(async () => ({ success: true, data: { stdout: "ok" } }));
    const executors = new Map<string, ToolExecutor>([
      ["shell.execute", { name: "shell.execute", execute: executor }],
    ]);
    const requestToolApproval = vi.fn(async () => ({ approved: false }));
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const sessionStoreAppend = vi.fn(async () => {});

    await processToolCalls(
      [{ id: "call-nohup", name: "shell.execute", arguments: "{\"command\":\"nohup node server.js\"}" }],
      new Map([["call-nohup", activeItem]]),
      turn,
      executors,
      { permissionMode: "confirm_all", requestToolApproval },
      callbacks,
      sessionStoreAppend
    );

    expect(requestToolApproval).toHaveBeenCalledWith(expect.objectContaining({
      toolName: "shell.execute",
      sandboxJustification: expect.stringContaining("后台长期运行"),
    }));
    expect(executor).not.toHaveBeenCalled();
  });
});

// ============================================================
// AgentLoop mid-turn compaction state
// ============================================================

describe("AgentLoop mid-turn compaction", () => {
  beforeEach(() => {
    aiClientMocks.streamChatParams.length = 0;
    aiClientMocks.client.streamChat.mockReset();
    aiClientMocks.client.chat.mockReset();
  });

  it("reports context usage from the estimated request payload", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "done" },
        { type: "done", finishReason: "stop" },
      ]);
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
          contextWindowSize: 100_000,
        },
        systemPrompt: "hidden request context ".repeat(250),
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 80_000,
          retainedUserMessageMaxTokens: 20_000,
          contextWindowSize: 100_000,
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          ["range.read", { name: "range.read", execute: vi.fn(async () => ({ success: true })) }],
        ]),
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "hi" }, callbacks);

    expect(aiClientMocks.streamChatParams[0].tools.length).toBeGreaterThan(0);
    const usage = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "context_usage"
    )?.[0];
    expect(usage).toMatchObject({
      type: "context_usage",
      contextWindowSize: 100_000,
    });
    expect(usage?.type === "context_usage" ? usage.estimatedTokens : 0).toBeGreaterThan(1000);
  });

  it("keeps the current user message visible after mid-turn compaction", async () => {
    const userContent = "请读取当前表格并继续分析";
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        return streamEvents([
          { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
          { type: "tool_call_end", toolCallId: "call-1", toolName: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" },
          { type: "done", finishReason: "tool_calls" },
        ]);
      }
      return streamEvents([
        { type: "text_delta", delta: "继续分析完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat.mockResolvedValue({ content: "压缩后的工具执行摘要" });

    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 1,
          retainedUserMessageMaxTokens: 0,
          contextWindowSize: 100,
          summaryRetryBaseDelayMs: 0,
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          [
            "range.read",
            {
              name: "range.read",
              execute: vi.fn(async () => ({ success: true, data: "A1=42" })),
            },
          ],
        ]),
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: userContent }, { onEvent: vi.fn() });

    expect(aiClientMocks.streamChatParams).toHaveLength(2);
    const secondMessages = aiClientMocks.streamChatParams[1].messages;
    expect(secondMessages.at(-1)).toMatchObject({
      role: "user",
      content: userContent,
    });
  });

  it("retries transient sampling request failures before failing the turn", async () => {
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        const error = new Error("temporary upstream failure");
        (error as any).status = 500;
        throw error;
      }
      return streamEvents([
        { type: "text_delta", delta: "重试后完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });

    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        aiRequestRetryConfig: {
          sampling: { maxRetries: 1, baseDelayMs: 0, maxDelayMs: 0 },
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "触发采样重试" }, { onEvent: vi.fn() });

    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(2);
    expect(aiClientMocks.streamChatParams).toHaveLength(2);
  });

  it("does not retry aborted sampling requests", async () => {
    const abortError = new Error("aborted");
    abortError.name = "AbortError";
    aiClientMocks.client.streamChat.mockImplementation(() => {
      throw abortError;
    });

    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        aiRequestRetryConfig: {
          sampling: { maxRetries: 2, baseDelayMs: 0, maxDelayMs: 0 },
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await expect(
      loop.runTurn({ content: "中断不应重试" }, { onEvent: vi.fn() })
    ).rejects.toBe(abortError);
    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(1);
  });

  it("uses compact retry backoff settings from compaction config", async () => {
    vi.useFakeTimers();
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        return streamEvents([
          { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
          { type: "tool_call_end", toolCallId: "call-1", toolName: "range.read", arguments: "{}" },
          { type: "done", finishReason: "tool_calls" },
        ]);
      }
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    let resolveFirstChat!: () => void;
    const firstChatCalled = new Promise<void>((resolve) => {
      resolveFirstChat = resolve;
    });
    aiClientMocks.client.chat
      .mockImplementationOnce(async () => {
        resolveFirstChat();
        throw Object.assign(new Error("rate limited"), { status: 429 });
      })
      .mockResolvedValueOnce({ content: "压缩摘要" });

    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 1,
          retainedUserMessageMaxTokens: 0,
          contextWindowSize: 100,
          summaryRetryCount: 1,
          summaryRetryBaseDelayMs: 50,
          summaryRetryMaxDelayMs: 50,
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          ["range.read", { name: "range.read", execute: vi.fn(async () => ({ success: true, data: "ok" })) }],
        ]),
      },
      createMemorySessionStore()
    );

    try {
      const runPromise = loop.runTurn({ content: "触发压缩 backoff" }, { onEvent: vi.fn() });
      await firstChatCalled;
      expect(aiClientMocks.client.chat).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(49);
      expect(aiClientMocks.client.chat).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(1);
      await runPromise;
      expect(aiClientMocks.client.chat).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("emits visible compaction progress items when mid-turn compaction runs", async () => {
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        return streamEvents([
          { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
          { type: "tool_call_end", toolCallId: "call-1", toolName: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" },
          { type: "done", finishReason: "tool_calls" },
        ]);
      }
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat.mockResolvedValue({ content: "压缩摘要" });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const sessionStore = createMemorySessionStore();
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 1,
          retainedUserMessageMaxTokens: 0,
          contextWindowSize: 100,
          summaryRetryBaseDelayMs: 0,
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          [
            "range.read",
            {
              name: "range.read",
              execute: vi.fn(async () => ({ success: true, data: "A1=42" })),
            },
          ],
        ]),
      },
      sessionStore
    );

    await loop.runTurn({ content: "触发压缩" }, callbacks);

    const started = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "item_started" && event.item.type === "compact_progress"
    )?.[0];
    const completed = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "item_completed" && event.item.type === "compact_progress"
    )?.[0];
    const compactStarted = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "thread_compact_started"
    )?.[0];

    expect(compactStarted).toMatchObject({
      type: "thread_compact_started",
      params: {
        reason: "auto_token_limit",
        itemCount: expect.any(Number),
        tokensBefore: expect.any(Number),
        retryCount: expect.any(Number),
      },
    });
    expect(started).toMatchObject({
      type: "item_started",
      item: {
        type: "compact_progress",
        status: "running",
        reason: "auto_token_limit",
      },
    });
    expect(completed).toMatchObject({
      type: "item_completed",
      item: {
        type: "compact_progress",
        status: "completed",
        summary: "压缩摘要",
        tokensBefore: expect.any(Number),
        tokensAfter: expect.any(Number),
      },
    });
    expect(sessionStore.appendRolloutItems).toHaveBeenCalledWith(
      expect.any(String),
      expect.arrayContaining([
        expect.objectContaining({
          type: "compact_params",
          reason: "auto_token_limit",
          status: "started",
        }),
      ])
    );
  });

  it("retries summary generation when compaction fails once", async () => {
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        return streamEvents([
          { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
          { type: "tool_call_end", toolCallId: "call-1", toolName: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" },
          { type: "done", finishReason: "tool_calls" },
        ]);
      }
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat
      .mockRejectedValueOnce(Object.assign(new Error("temporary summary failure"), { status: 500 }))
      .mockResolvedValueOnce({ content: "重试后的压缩摘要" });

    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 1,
          retainedUserMessageMaxTokens: 0,
          contextWindowSize: 100,
          summaryRetryBaseDelayMs: 0,
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          [
            "range.read",
            {
              name: "range.read",
              execute: vi.fn(async () => ({ success: true, data: "A1=42" })),
            },
          ],
        ]),
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "触发压缩重试" }, { onEvent: vi.fn() });

    expect(aiClientMocks.client.chat).toHaveBeenCalledTimes(2);
  });

  it("compacts before the next turn after the model changes", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat.mockResolvedValue({ content: "模型切换摘要" });
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "model-a",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 100_000,
          retainedUserMessageMaxTokens: 20_000,
          contextWindowSize: 100_000,
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "先建立历史" }, callbacks);
    loop.updateAIConfig({
      provider: "openai",
      apiKey: "test",
      baseUrl: "https://example.test",
      model: "model-b",
      reasoningMode: "off",
    });
    await loop.runTurn({ content: "模型切换后继续" }, callbacks);

    const compactStarted = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "thread_compact_started"
        && event.params.reason === "model_changed"
    );
    expect(compactStarted).toBeDefined();
    expect(aiClientMocks.client.chat).toHaveBeenCalledWith(expect.objectContaining({
      messages: expect.arrayContaining([
        expect.objectContaining({ content: expect.stringContaining("先建立历史") }),
      ]),
    }));
  });

  it("skips model-change compaction when compHash is unchanged", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "model-a",
          compHash: "office-chat-v1",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 100_000,
          retainedUserMessageMaxTokens: 20_000,
          contextWindowSize: 100_000,
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "先建立历史" }, callbacks);
    loop.updateAIConfig({
      provider: "custom",
      apiKey: "test",
      baseUrl: "https://example.test",
      model: "model-b",
      compHash: "office-chat-v1",
      reasoningMode: "off",
    });
    await loop.runTurn({ content: "同兼容族继续" }, callbacks);

    const compactStarted = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "thread_compact_started"
        && event.params.reason === "model_changed"
    );
    expect(compactStarted).toBeUndefined();
    expect(aiClientMocks.client.chat).not.toHaveBeenCalled();
  });

  it("compacts when compHash changes even if provider and model stay the same", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat.mockResolvedValue({ content: "兼容性变化摘要" });
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "model-a",
          compHash: "chat-v1",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 100_000,
          retainedUserMessageMaxTokens: 20_000,
          contextWindowSize: 100_000,
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "先建立历史" }, callbacks);
    loop.updateAIConfig({
      provider: "openai",
      apiKey: "test",
      baseUrl: "https://example.test",
      model: "model-a",
      compHash: "responses-v2",
      reasoningMode: "off",
    });
    await loop.runTurn({ content: "兼容性变化后继续" }, callbacks);

    const compactStarted = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "thread_compact_started"
        && event.params.reason === "model_changed"
    );
    expect(compactStarted).toBeDefined();
  });

  it("compacts before the next turn after the context window changes", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    aiClientMocks.client.chat.mockResolvedValue({ content: "窗口变化摘要" });
    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "model-a",
          reasoningMode: "off",
        },
        compactionConfig: {
          enabled: true,
          autoCompactTokenThreshold: 100_000,
          retainedUserMessageMaxTokens: 20_000,
          contextWindowSize: 100_000,
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    await loop.runTurn({ content: "先建立历史" }, callbacks);
    loop.updateCompactionConfig({
      enabled: true,
      autoCompactTokenThreshold: 40_000,
      retainedUserMessageMaxTokens: 20_000,
      contextWindowSize: 50_000,
    });
    await loop.runTurn({ content: "窗口变化后继续" }, callbacks);

    const compactStarted = vi.mocked(callbacks.onEvent).mock.calls.find(
      ([event]) => event.type === "thread_compact_started"
        && event.params.reason === "context_window_changed"
    );
    expect(compactStarted).toBeDefined();
  });
});

// ============================================================
// AgentLoop thread idle unload
// ============================================================

describe("AgentLoop thread idle unload", () => {
  beforeEach(() => {
    aiClientMocks.streamChatParams.length = 0;
    aiClientMocks.client.streamChat.mockReset();
    aiClientMocks.client.chat.mockReset();
  });

  it("unloads idle thread memory while keeping the thread resumable", async () => {
    const sessionStore = createMemorySessionStore();
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
        threadIdleUnloadMs: 1_000,
      },
      sessionStore
    );

    const threadId = await loop.startThread();

    expect(loop.getThread()?.metadata.threadId).toBe(threadId);

    const unloaded = await loop.sweepIdleThread(Number.MAX_SAFE_INTEGER);

    expect(unloaded).toBe(true);
    expect(loop.getThread()).toBeNull();
    expect(loop.getThreadRuntimeStatus()).toMatchObject({
      status: "unloaded",
      threadId,
    });

    await expect(loop.resumeThread(threadId)).resolves.toBe(true);
    expect(loop.getThread()?.metadata.threadId).toBe(threadId);
  });

  it("persists thread snapshots and runtime state when threads start and unload", async () => {
    const sessionStore = createMemorySessionStore();
    const stateRuntimeStore = {
      upsertThreadSnapshot: vi.fn(async () => {}),
      updateThreadRuntime: vi.fn(async () => {}),
    };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
        threadIdleUnloadMs: 1_000,
      },
      sessionStore,
      stateRuntimeStore as any
    );

    const threadId = await loop.startThread();

    expect(stateRuntimeStore.upsertThreadSnapshot).toHaveBeenCalledWith(
      expect.objectContaining({ threadId })
    );
    expect(stateRuntimeStore.updateThreadRuntime).toHaveBeenLastCalledWith(
      expect.objectContaining({ threadId, status: "active" })
    );

    await loop.sweepIdleThread(Number.MAX_SAFE_INTEGER);

    expect(stateRuntimeStore.updateThreadRuntime).toHaveBeenLastCalledWith(
      expect.objectContaining({ threadId, status: "unloaded" })
    );
  });

  it("publishes thread runtime status to subscribed connections", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "text_delta", delta: "完成" },
        { type: "done", finishReason: "stop" },
      ]);
    });
    const watchManager = new ThreadWatchManager();
    const sessionStore = createMemorySessionStore();
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
      },
      sessionStore,
      undefined,
      watchManager
    );

    const threadId = await loop.startThread();
    const statuses: string[] = [];
    watchManager.watch(threadId, "connection-1", (status) => {
      statuses.push(status.status);
    });

    await loop.runTurn({ content: "观察线程状态" }, { onEvent: vi.fn() });

    expect(statuses).toEqual(["active", "running", "active"]);
    expect(watchManager.getConnectionIds(threadId)).toEqual(["connection-1"]);
  });
});

// ============================================================
// AgentLoop input queue
// ============================================================

describe("AgentLoop input queue", () => {
  beforeEach(() => {
    aiClientMocks.streamChatParams.length = 0;
    aiClientMocks.client.streamChat.mockReset();
    aiClientMocks.client.chat.mockReset();
  });

  it("queues user input while a turn is running and drains it after the current turn completes", async () => {
    let releaseFirstTurn!: () => void;
    const firstTurnCanFinish = new Promise<void>((resolve) => {
      releaseFirstTurn = resolve;
    });
    let streamCallCount = 0;
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      streamCallCount += 1;
      if (streamCallCount === 1) {
        return (async function* (): AsyncIterable<AIStreamEvent> {
          yield { type: "text_delta", delta: "正在分析 Sheet1" };
          await firstTurnCanFinish;
          yield { type: "done", finishReason: "stop" };
        })();
      }
      return streamEvents([
        { type: "text_delta", delta: "已改用 Sheet2" },
        { type: "done", finishReason: "stop" },
      ]);
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    const firstRun = loop.runTurn({ content: "先分析 Sheet1" }, callbacks);
    await waitForEvent(callbacks, "turn_started");

    const queued = loop.enqueueTurn({ content: "等一下，用 Sheet2 的数据" }, callbacks);

    expect(queued).toEqual({ queued: true, queueSize: 1 });
    expect(loop.getQueuedInputCount()).toBe(1);

    releaseFirstTurn();
    await firstRun;
    await waitForEventCount(callbacks, "turn_completed", 2);

    expect(loop.getQueuedInputCount()).toBe(0);
    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(2);
    expect(aiClientMocks.streamChatParams[1].messages.at(-1)).toMatchObject({
      role: "user",
      content: "等一下，用 Sheet2 的数据",
    });
  });

  it("clears queued user input when the active turn is interrupted", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return (async function* (): AsyncIterable<AIStreamEvent> {
        await new Promise<void>((_resolve, reject) => {
          params.signal?.addEventListener("abort", () => {
            const error = new Error("aborted");
            error.name = "AbortError";
            reject(error);
          });
        });
      })();
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    const runResult = loop.runTurn({ content: "先保持运行" }, callbacks).catch((error) => error);
    await waitForEvent(callbacks, "turn_started");

    loop.enqueueTurn({ content: "这条会被中断清理" }, callbacks);
    expect(loop.getQueuedInputCount()).toBe(1);

    await loop.interrupt("request-clear-queue");
    await expect(runResult).resolves.toMatchObject({ name: "AbortError" });

    expect(loop.getQueuedInputCount()).toBe(0);
    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(1);
  });

  it("does not auto-drain input queued during an interrupt race", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return (async function* (): AsyncIterable<AIStreamEvent> {
        await new Promise<void>((_resolve, reject) => {
          params.signal?.addEventListener("abort", () => {
            setTimeout(() => {
              const error = new Error("aborted");
              error.name = "AbortError";
              reject(error);
            }, 0);
          });
        });
      })();
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
      },
      createMemorySessionStore()
    );

    const runResult = loop.runTurn({ content: "运行后中断" }, callbacks).catch((error) => error);
    await waitForEvent(callbacks, "turn_started");

    const interruptResult = loop.interrupt("request-race");
    expect(() => loop.enqueueTurn({ content: "不应该自动恢复" }, callbacks)).toThrow("Agent 正在中断中");

    await interruptResult;
    await expect(runResult).resolves.toMatchObject({ name: "AbortError" });
    await new Promise((resolve) => setTimeout(resolve, 5));

    expect(loop.getQueuedInputCount()).toBe(0);
    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(1);
  });

  it("stops after an in-flight tool finishes when interrupted during tool execution", async () => {
    let releaseTool!: () => void;
    const toolCanFinish = new Promise<void>((resolve) => {
      releaseTool = resolve;
    });
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return streamEvents([
        { type: "tool_call_begin", toolCallId: "call-slow", toolName: "range.read" },
        { type: "tool_call_end", toolCallId: "call-slow", toolName: "range.read", arguments: "{}" },
        { type: "done", finishReason: "tool_calls" },
      ]);
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>([
          ["range.read", {
            name: "range.read",
            execute: vi.fn(async () => {
              await toolCanFinish;
              return { success: true, data: "ok" };
            }),
          }],
        ]),
      },
      createMemorySessionStore()
    );

    const runResult = loop.runTurn({ content: "执行慢工具后停止" }, callbacks).catch((error) => error);
    await waitForEvent(callbacks, "turn_started");
    await waitForEvent(callbacks, "item_started");

    const interruptResult = loop.interrupt("request-during-tool");
    releaseTool();

    await interruptResult;
    await expect(runResult).resolves.toMatchObject({ name: "AbortError" });
    expect(aiClientMocks.client.streamChat).toHaveBeenCalledTimes(1);
    expect(vi.mocked(callbacks.onEvent).mock.calls.some(
      ([event]) => event.type === "turn_interrupted"
    )).toBe(true);
  });
});

// ============================================================
// AgentLoop interrupt queue
// ============================================================

describe("AgentLoop interrupt queue", () => {
  beforeEach(() => {
    aiClientMocks.streamChatParams.length = 0;
    aiClientMocks.client.streamChat.mockReset();
    aiClientMocks.client.chat.mockReset();
  });

  it("tracks pending interrupt request ids until the active turn is cleaned up", async () => {
    aiClientMocks.client.streamChat.mockImplementation((params: StreamChatParams) => {
      aiClientMocks.streamChatParams.push(params);
      return (async function* (): AsyncIterable<AIStreamEvent> {
        await new Promise<void>((_resolve, reject) => {
          params.signal?.addEventListener("abort", () => {
            const error = new Error("aborted");
            error.name = "AbortError";
            reject(error);
          });
        });
      })();
    });

    const callbacks: AgentTurnCallbacks = { onEvent: vi.fn() };
    const loop = new AgentLoop(
      {
        aiConfig: {
          provider: "openai",
          apiKey: "test",
          baseUrl: "https://example.test",
          model: "test-model",
          reasoningMode: "off",
        },
        permissionMode: "confirm_all",
        toolExecutors: new Map<string, ToolExecutor>(),
      },
      createMemorySessionStore()
    );

    const runResult = loop.runTurn({ content: "开始后等待中断" }, callbacks).catch((error) => error);
    await waitForEvent(callbacks, "turn_started");

    const interruptResult = loop.interrupt("request-1");
    expect(loop.getPendingInterruptRequestIds()).toEqual(["request-1"]);

    await interruptResult;
    await expect(runResult).resolves.toMatchObject({ name: "AbortError" });
    expect(loop.getPendingInterruptRequestIds()).toEqual([]);
  });
});

async function waitForEvent(callbacks: AgentTurnCallbacks, eventType: string): Promise<void> {
  for (let i = 0; i < 20; i++) {
    if (vi.mocked(callbacks.onEvent).mock.calls.some(([event]) => event.type === eventType)) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`event not received: ${eventType}`);
}

async function waitForEventCount(
  callbacks: AgentTurnCallbacks,
  eventType: string,
  count: number
): Promise<void> {
  for (let i = 0; i < 40; i++) {
    const actual = vi.mocked(callbacks.onEvent).mock.calls.filter(
      ([event]) => event.type === eventType
    ).length;
    if (actual >= count) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`event count not received: ${eventType} x ${count}`);
}

// ============================================================
// shouldRequireApproval
// ============================================================

describe("shouldRequireApproval", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools();
  });

  // ---- normal 模式：所有工具都需要审批（但 alwaysAllowedTools 可跳过） ----
  it("normal: all tools require approval unless alwaysAllowedTools", () => {
    expect(shouldRequireApproval("shell.execute", "normal")).toBe(true);
    expect(shouldRequireApproval("range.read", "normal")).toBe(true);
    expect(shouldRequireApproval("range.write", "normal")).toBe(true);
    expect(shouldRequireApproval("range.clear", "normal")).toBe(true);
    expect(shouldRequireApproval("workbook.inspect", "normal")).toBe(true);
    expect(shouldRequireApproval("unknown_tool", "normal")).toBe(true);
  });

  it("normal: alwaysAllowedTools overrides approval", () => {
    markToolAlwaysAllowed("shell.execute");
    expect(shouldRequireApproval("shell.execute", "normal")).toBe(false);
    clearAlwaysAllowedTools();
    markToolAlwaysAllowed("range.write");
    expect(shouldRequireApproval("range.write", "normal")).toBe(false);
  });

  // ---- auto_approve_safe：safe 级别自动批准 ----
  it("auto_approve_safe: safe tools are auto-approved", () => {
    expect(shouldRequireApproval("range.read", "auto_approve_safe")).toBe(false);
    expect(shouldRequireApproval("workbook.inspect", "auto_approve_safe")).toBe(false);
    expect(shouldRequireApproval("selection.get", "auto_approve_safe")).toBe(false);
  });

  it("auto_approve_safe: moderate and dangerous tools require approval", () => {
    expect(shouldRequireApproval("range.write", "auto_approve_safe")).toBe(true);
    expect(shouldRequireApproval("range.clear", "auto_approve_safe")).toBe(true);
    expect(shouldRequireApproval("shell.execute", "auto_approve_safe")).toBe(true);
    expect(shouldRequireApproval("macro.write", "auto_approve_safe")).toBe(true);
  });

  it("auto_approve_safe: unknown tools default to requiring approval", () => {
    expect(shouldRequireApproval("unknown_tool", "auto_approve_safe")).toBe(true);
  });

  // ---- confirm_all：全部确认，所有工具自动批准 ----
  it("confirm_all: all tools including file deletions are auto-approved", () => {
    expect(shouldRequireApproval("range.clear", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("sheet.operation", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("ui.removeControl", "confirm_all")).toBe(false);
  });

  it("confirm_all: non-deletion tools are auto-approved", () => {
    expect(shouldRequireApproval("range.read", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("range.write", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("shell.execute", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("macro.write", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("macro.run", "confirm_all")).toBe(false);
    expect(shouldRequireApproval("workbook.save", "confirm_all")).toBe(false);
  });

  it("confirm_all: unknown tools default to auto-approved", () => {
    expect(shouldRequireApproval("unknown_tool", "confirm_all")).toBe(false);
  });

  // ---- alwaysAllowedTools 行为 ----
  it("alwaysAllowedTools: overrides auto_approve_safe and confirm_all", () => {
    markToolAlwaysAllowed("range.clear");
    expect(shouldRequireApproval("range.clear", "auto_approve_safe")).toBe(false);
    expect(shouldRequireApproval("range.clear", "confirm_all")).toBe(false);
  });
});

// ============================================================
// markToolAlwaysAllowed / getAlwaysAllowedTools
// ============================================================

describe("alwaysAllowedTools", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools();
  });

  it("should track always-allowed tools", () => {
    markToolAlwaysAllowed("tool_a");
    markToolAlwaysAllowed("tool_b");
    const tools = getAlwaysAllowedTools();
    expect(tools.has("tool_a")).toBe(true);
    expect(tools.has("tool_b")).toBe(true);
  });

  it("should clear all", () => {
    markToolAlwaysAllowed("tool_a");
    clearAlwaysAllowedTools();
    expect(getAlwaysAllowedTools().size).toBe(0);
  });
});
