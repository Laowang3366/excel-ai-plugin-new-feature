import { mkdtemp, rm } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it, vi } from "vitest";

import { SessionStore } from "./sessionStore";
import { StateRuntimeStore } from "./stateRuntimeStore";
import { ThreadRepository } from "./threadRepository";

describe("ThreadRepository", () => {
  it("keeps JSONL metadata and sqlite snapshots consistent across update and delete", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "thread-repository-"));
    const sessions = new SessionStore(path.join(root, "sessions"));
    const runtime = new StateRuntimeStore(path.join(root, "runtime"));
    await runtime.init();

    try {
      const repository = new ThreadRepository(sessions, runtime);
      const thread = await sessions.createThread("openai", "model-a", "folder-old");
      await runtime.upsertThreadSnapshot(thread.metadata);
      const createdAt = thread.metadata.createdAt;

      await repository.updateMetadata(thread.metadata.threadId, {
        name: "季度分析",
        folderId: "folder-new",
      });

      const loaded = await sessions.loadThread(thread.metadata.threadId);
      const snapshot = await runtime.getThreadSnapshot(thread.metadata.threadId);
      expect(loaded?.metadata).toMatchObject({
        name: "季度分析",
        folderId: "folder-new",
        createdAt,
      });
      expect(snapshot).toMatchObject({ name: "季度分析", folderId: "folder-new" });

      await repository.updateMetadata(thread.metadata.threadId, { name: "", folderId: undefined });
      expect(
        (await sessions.loadThread(thread.metadata.threadId))?.metadata.folderId,
      ).toBeUndefined();
      expect((await runtime.getThreadSnapshot(thread.metadata.threadId))?.folderId).toBeUndefined();
      await expect(runtime.findThreadNameByIdStr(thread.metadata.threadId)).resolves.toBeNull();

      await expect(repository.delete(thread.metadata.threadId)).resolves.toBe(true);
      await expect(sessions.loadThread(thread.metadata.threadId)).resolves.toBeNull();
      await expect(runtime.getThreadSnapshot(thread.metadata.threadId)).resolves.toBeNull();
      await expect(runtime.listRolloutEvents(thread.metadata.threadId)).resolves.toEqual([]);
    } finally {
      await runtime.close();
      await rm(root, { recursive: true, force: true });
    }
  });

  it("backfills legacy JSONL threads when sqlite snapshots are empty", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "thread-repository-"));
    const sessions = new SessionStore(path.join(root, "sessions"));
    const runtime = new StateRuntimeStore(path.join(root, "runtime"));
    await runtime.init();

    try {
      const thread = await sessions.createThread("openai", "model-a");
      const repository = new ThreadRepository(sessions, runtime);
      const listed = await repository.list();

      expect(listed.map((item) => item.threadId)).toEqual([thread.metadata.threadId]);
      expect(await runtime.getThreadSnapshot(thread.metadata.threadId)).toMatchObject({
        threadId: thread.metadata.threadId,
      });
    } finally {
      await runtime.close();
      await rm(root, { recursive: true, force: true });
    }
  });

  it("keeps sqlite projections when rollout artifact deletion fails", async () => {
    const sessions = {
      deleteThread: vi.fn(async () => {
        throw new Error("archive delete denied");
      }),
    } as unknown as SessionStore;
    const runtime = {
      getThreadSnapshot: vi.fn(async () => ({ threadId: "thread-protected" })),
      deleteThreadData: vi.fn(async () => undefined),
    } as unknown as StateRuntimeStore;
    const repository = new ThreadRepository(sessions, runtime);

    await expect(repository.delete("thread-protected")).rejects.toThrow("archive delete denied");
    expect(runtime.deleteThreadData).not.toHaveBeenCalled();
  });
});
