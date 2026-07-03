import type { AIClientConfig } from "../../providers/aiClient";
import { SessionStore } from "../../memory/sessionStore";
import { resolveModelCompHash } from "./modelCompHash";
import {
  DEFAULT_COMPACTION_CONFIG,
  type CompactionConfig,
  type Thread,
  type ThreadId,
  type TurnItem,
} from "../../shared/types";

export interface StartThreadParams {
  sessionStore: SessionStore;
  aiConfig: AIClientConfig;
  compactionConfig?: CompactionConfig;
  folderId?: string;
}

export interface ResumeThreadResult {
  thread: Thread;
  compactedHistory: TurnItem[] | null;
}

/**
 * 线程生命周期操作。
 *
 * 关联模块：
 * - memory/sessionStore: 创建、加载、注册 rollout 路径。
 * - agentLoop.ts: 保留公共 API，委托这里处理线程持久化细节。
 */
export async function createAgentThread(params: StartThreadParams): Promise<Thread> {
  const thread = await params.sessionStore.createThread(
    params.aiConfig.provider,
    params.aiConfig.model,
    params.folderId
  );
  const compactionConfig = params.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
  thread.metadata.contextWindowSize = compactionConfig.contextWindowSize || params.aiConfig.contextWindowSize || 128_000;
  thread.metadata.compHash = resolveModelCompHash(params.aiConfig);
  return thread;
}

export async function loadAgentThread(
  sessionStore: SessionStore,
  threadId: ThreadId
): Promise<ResumeThreadResult | null> {
  const thread = await sessionStore.loadThread(threadId);
  if (!thread) return null;

  const rolloutPath = await sessionStore.findRolloutPath(threadId);
  if (rolloutPath) {
    sessionStore.registerRolloutPath(threadId, rolloutPath);
  }

  return {
    thread,
    compactedHistory: thread.metadata.compactedHistory ?? null,
  };
}
