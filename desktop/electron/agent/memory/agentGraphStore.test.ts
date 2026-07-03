import { mkdtemp, rm } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it } from "vitest";

import { AgentGraphStore } from "./agentGraphStore";

describe("AgentGraphStore", () => {
  it("persists thread spawn edges and lists descendants breadth-first", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "agent-graph-store-"));
    try {
      const store = new AgentGraphStore(tempDir);

      await store.upsertThreadSpawnEdge("thread-root", "thread-child-a", { createdAt: 1 });
      await store.upsertThreadSpawnEdge("thread-root", "thread-child-b", { createdAt: 2 });
      await store.upsertThreadSpawnEdge("thread-child-a", "thread-grandchild", { createdAt: 3 });

      const reloaded = new AgentGraphStore(tempDir);
      const descendants = await reloaded.listThreadSpawnDescendants("thread-root");

      expect(descendants.map((item) => item.threadId)).toEqual([
        "thread-child-a",
        "thread-child-b",
        "thread-grandchild",
      ]);
      expect(descendants.map((item) => item.depth)).toEqual([1, 1, 2]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("filters descendants by edge status", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "agent-graph-store-"));
    try {
      const store = new AgentGraphStore(tempDir);

      await store.upsertThreadSpawnEdge("thread-root", "thread-open");
      await store.upsertThreadSpawnEdge("thread-root", "thread-closed");
      await store.closeThreadSpawnEdge("thread-root", "thread-closed", 10);

      const openDescendants = await store.listThreadSpawnDescendants("thread-root", {
        status: "open",
      });
      const closedDescendants = await store.listThreadSpawnDescendants("thread-root", {
        status: "closed",
      });

      expect(openDescendants.map((item) => item.threadId)).toEqual(["thread-open"]);
      expect(closedDescendants.map((item) => item.threadId)).toEqual(["thread-closed"]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
