import { describe, expect, it } from "vitest";
import { getSidebarThreadItemStatus } from "./SidebarThreadItem";

const baseThread = {
  threadId: "thread-1",
  preview: "Quarterly report",
  updatedAt: 100,
};

describe("getSidebarThreadItemStatus", () => {
  it("marks running threads from runtime state", () => {
    expect(getSidebarThreadItemStatus({
      thread: baseThread,
      activeThreadId: null,
      runningThreadIds: { "thread-1": true },
      turnStatus: "idle",
      viewedThreadStatusAt: {},
    })).toEqual({ isActiveThread: false, status: "running" });
  });

  it("marks the active thread as viewed for non-running status badges", () => {
    expect(getSidebarThreadItemStatus({
      thread: baseThread,
      activeThreadId: "thread-1",
      runningThreadIds: {},
      turnStatus: "failed",
      viewedThreadStatusAt: {},
    })).toEqual({ isActiveThread: true, status: null });
  });

  it("hides completed metadata after the user has viewed it", () => {
    expect(getSidebarThreadItemStatus({
      thread: { ...baseThread, lastTurnStatus: "completed" },
      activeThreadId: null,
      runningThreadIds: {},
      turnStatus: "idle",
      viewedThreadStatusAt: { "thread-1": 100 },
    })).toEqual({ isActiveThread: false, status: null });
  });
});
