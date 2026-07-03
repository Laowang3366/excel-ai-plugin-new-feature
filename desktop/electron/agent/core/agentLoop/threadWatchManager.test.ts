import { describe, expect, it } from "vitest";

import { ThreadWatchManager } from "./threadWatchManager";

describe("ThreadWatchManager", () => {
  it("tracks watcher connection ids and publishes thread status changes", () => {
    const manager = new ThreadWatchManager();
    const statuses: string[] = [];

    const watcher = manager.watch("thread-1", "connection-1", (status) => {
      statuses.push(status.status);
    });

    manager.publish({
      threadId: "thread-1",
      status: "active",
      idleUnloadMs: 1000,
    });

    expect(statuses).toEqual(["active"]);
    expect(manager.getConnectionIds("thread-1")).toEqual(["connection-1"]);

    watcher.close();

    expect(manager.getConnectionIds("thread-1")).toEqual([]);
  });

  it("uses an active guard to release observed activity when disposed", () => {
    const manager = new ThreadWatchManager();
    const guard = manager.createActiveGuard("thread-1", "connection-1");

    expect(manager.getActiveConnectionIds("thread-1")).toEqual(["connection-1"]);

    guard.dispose();
    guard.dispose();

    expect(manager.getActiveConnectionIds("thread-1")).toEqual([]);
  });
});
