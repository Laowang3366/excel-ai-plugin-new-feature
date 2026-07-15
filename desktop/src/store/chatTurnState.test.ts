import { describe, expect, it } from "vitest";
import { buildTurnStartPatch } from "./chatTurnState";

describe("buildTurnStartPatch", () => {
  it("clears streaming state and removes the active thread from stopped thread ids", () => {
    const patch = buildTurnStartPatch(
      {
        activeThreadId: "thread-active",
        stoppedThreadIds: {
          "thread-active": true,
          "thread-other": true,
        },
      },
      "client-1",
      { compactionNotice: null },
    );

    expect(patch).toEqual({
      isStreaming: true,
      streamingContent: "",
      streamingReasoning: "",
      activeStreamingRound: null,
      turnStatus: "in_progress",
      activeClientId: "client-1",
      stoppedThreadIds: {
        "thread-other": true,
      },
      error: null,
      compactionNotice: null,
    });
  });

  it("keeps stopped thread ids unchanged when there is no active thread", () => {
    const stoppedThreadIds = { "thread-stopped": true };

    const patch = buildTurnStartPatch(
      {
        activeThreadId: null,
        stoppedThreadIds,
      },
      "client-2",
      { lastInterruptContext: null },
    );

    expect(patch.stoppedThreadIds).toBe(stoppedThreadIds);
    expect(patch.lastInterruptContext).toBeNull();
  });
});
