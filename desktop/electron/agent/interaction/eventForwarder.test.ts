import { afterEach, describe, expect, it, vi } from "vitest";

// @MOCK_INTERFACE: provides the Electron BrowserWindow/ipcMain surface consumed by eventForwarder.
vi.mock("electron", () => ({
  BrowserWindow: class {},
  ipcMain: { handle: vi.fn() },
}));

import {
  STREAM_DELTA_FORWARD_FLUSH_MS,
  createEventForwarder,
  mergeStreamDeltas,
} from "./eventForwarder";

function createMockWindow() {
  // @MOCK_INTERFACE: minimal BrowserWindow instance contract used for agent event forwarding.
  return {
    isDestroyed: vi.fn(() => false),
    webContents: {
      send: vi.fn(),
    },
  };
}

describe("eventForwarder stream delta batching", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.clearAllMocks();
  });

  it("merges adjacent stream deltas with the same stream identity", () => {
    expect(
      mergeStreamDeltas([
        { delta: "思", itemType: "reasoning", roundId: 1, threadId: "t" },
        { delta: "考", itemType: "reasoning", roundId: 1, threadId: "t" },
        { delta: "正文", itemType: "assistant_message", roundId: 1, threadId: "t" },
        { delta: "新轮", itemType: "reasoning", roundId: 2, threadId: "t" },
      ]),
    ).toEqual([
      { delta: "思考", itemType: "reasoning", roundId: 1, threadId: "t" },
      { delta: "正文", itemType: "assistant_message", roundId: 1, threadId: "t" },
      { delta: "新轮", itemType: "reasoning", roundId: 2, threadId: "t" },
    ]);
  });

  it("flushes pending stream deltas before forwarding an agent event", () => {
    vi.useFakeTimers();
    const mockWindow = createMockWindow();
    const forwarder = createEventForwarder(() => mockWindow as any);

    forwarder.onStreamDelta("思", "reasoning", 1, "thread-1", "client-1");
    forwarder.onStreamDelta("考", "reasoning", 1, "thread-1", "client-1");
    forwarder.onEvent({ type: "turn_started", turnId: "turn-1" } as any);

    expect(mockWindow.webContents.send).toHaveBeenNthCalledWith(1, "agent:event", {
      type: "stream_delta",
      delta: "思考",
      itemType: "reasoning",
      roundId: 1,
      threadId: "thread-1",
      clientId: "client-1",
    });
    expect(mockWindow.webContents.send).toHaveBeenNthCalledWith(2, "agent:event", {
      type: "turn_started",
      turnId: "turn-1",
    });
  });

  it("flushes batched stream deltas on the short timer", () => {
    vi.useFakeTimers();
    const mockWindow = createMockWindow();
    const forwarder = createEventForwarder(() => mockWindow as any);

    forwarder.onStreamDelta("A", "assistant_message", 1);
    forwarder.onStreamDelta("B", "assistant_message", 1);
    expect(mockWindow.webContents.send).not.toHaveBeenCalled();

    vi.advanceTimersByTime(STREAM_DELTA_FORWARD_FLUSH_MS);

    expect(mockWindow.webContents.send).toHaveBeenCalledWith("agent:event", {
      type: "stream_delta",
      delta: "AB",
      itemType: "assistant_message",
      roundId: 1,
      threadId: undefined,
      clientId: undefined,
    });
  });
});
