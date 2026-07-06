import { describe, expect, it } from "vitest";
import type { ThreadMetadata } from "../electronApi";
import { reconcileRunningThreadIds } from "./chatThreadRuntimeState";

describe("reconcileRunningThreadIds", () => {
  const thread = (threadId: string, patch: Partial<ThreadMetadata> = {}): ThreadMetadata => ({
    threadId,
    preview: threadId,
    modelProvider: "test",
    createdAt: 1,
    updatedAt: 1,
    ...patch,
  });

  it("keeps active or in-progress threads, removes listed completed entries, and preserves unlisted running ids", () => {
    expect(reconcileRunningThreadIds({
      threads: [
        thread("thread-active", { activeTurnId: "turn-1" }),
        thread("thread-progress", { lastTurnStatus: "in_progress" }),
        thread("thread-completed", { lastTurnStatus: "completed" }),
      ],
      runningThreadIds: {
        "thread-completed": true,
        "thread-stale": true,
      },
      stoppedThreadIds: {},
    })).toEqual({
      "thread-active": true,
      "thread-progress": true,
      "thread-stale": true,
    });
  });

  it("does not revive user-stopped in-progress metadata", () => {
    expect(reconcileRunningThreadIds({
      threads: [
        thread("thread-stopped", {
          activeTurnId: "turn-old",
          lastTurnStatus: "in_progress",
        }),
      ],
      runningThreadIds: {},
      stoppedThreadIds: {
        "thread-stopped": true,
      },
    })).toEqual({});
  });
});
