import type { ThreadId, ThreadMetadata } from "../shared/types";
import type { SessionStore } from "./sessionStore";
import type { StateRuntimeStore } from "./stateRuntimeStore";

export class ThreadRepository {
  constructor(
    private readonly sessions: SessionStore,
    private readonly runtime: StateRuntimeStore,
  ) {}

  async list(): Promise<ThreadMetadata[]> {
    const snapshots = await this.runtime.listThreadSnapshots();
    if (snapshots.length > 0) return snapshots;

    const legacyThreads = await this.sessions.listThreads();
    await Promise.all(legacyThreads.map((metadata) => this.runtime.upsertThreadSnapshot(metadata)));
    return legacyThreads;
  }

  async delete(threadId: ThreadId): Promise<boolean> {
    const hadSnapshot = Boolean(await this.runtime.getThreadSnapshot(threadId));
    const deletedSession = await this.sessions.deleteThread(threadId);
    await this.runtime.deleteThreadData(threadId);
    return deletedSession || hadSnapshot;
  }

  async updateMetadata(
    threadId: ThreadId,
    patch: Partial<ThreadMetadata>,
  ): Promise<ThreadMetadata> {
    await this.sessions.updateThreadMetadata(threadId, patch);
    const updated = await this.sessions.loadThread(threadId);
    if (!updated) throw new Error("会话不存在");
    await this.runtime.upsertThreadSnapshot(updated.metadata);
    return updated.metadata;
  }
}
