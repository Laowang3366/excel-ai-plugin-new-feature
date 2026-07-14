/**
 * Agent 相关 IPC 注册。
 *
 * 关联模块：
 * - interaction/eventForwarder: Agent 事件和工具审批 IPC。
 * - core/agentLoop: 启动、继续、中断、恢复线程。
 * - memory/sessionStore: 线程列表、元数据和统计。
 * - knowledge/knowledgeRegistry: RAG 知识库查询与索引入口。
 */

import { BrowserWindow } from "electron";
import { trustedIpcMain as ipcMain } from "../../shared/trustedIpc";
import type { AgentLoop } from "../core/agentLoop";
import type { AgentLoopManager } from "../runtime/agentRuntime";
import {
  getKnowledgeIndexer,
  getKnowledgeRetriever,
  getKnowledgeStore,
} from "../knowledge/knowledgeRegistry";
import type { KnowledgeRuntimeState } from "../runtime/knowledgeRuntime";
import type { SessionStore } from "../memory/sessionStore";
import type { AgentGraphStore } from "../memory/agentGraphStore";
import type { StateRuntimeStore } from "../memory/stateRuntimeStore";
import { ThreadRepository } from "../memory/threadRepository";
import { createLogger } from "../../shared/logger";
import type { AgentTurnCallbacks, AgentTurnInput, ThreadId } from "../shared/types";
import { ALL_TOOL_DEFINITIONS } from "../tools/registry/toolDefinitions";
import {
  AgentInterruptInput,
  AgentContinueTurnInput,
  AgentStartTurnInput,
  KnowledgeDeleteInput,
  KnowledgeIndexFileInput,
  KnowledgeIndexFolderInput,
  KnowledgeSearchInput,
  StatsGetSummaryInput,
  ThreadIdInput,
  ThreadGraphCloseEdgeInput,
  ThreadGraphEdgeInput,
  ThreadGraphListDescendantsInput,
  ThreadNewInput,
  ThreadUpdateMetadataInput,
  validateInput,
} from "../../shared/ipcSchemas";
import { createEventForwarder, registerToolApprovalHandlers } from "./eventForwarder";
import { assertAuthorizedPath, type PathAuthorizer } from "../../main-modules/ipcPathSecurity";
import type { IndexResult } from "../knowledge/types";

export interface AgentIpcHandlerDeps {
  mainWindowRef: () => BrowserWindow | null;
  agentLoopManagerRef: () => AgentLoopManager | null;
  getSessionStoreInstance: () => SessionStore;
  getStateRuntimeStoreInstance?: () => Promise<StateRuntimeStore>;
  getAgentGraphStoreInstance: () => AgentGraphStore;
  ensureKnowledgeRuntime?: () => Promise<KnowledgeRuntimeState>;
  isDataMigrationInProgress?: () => boolean;
  pathAuthorizer: PathAuthorizer;
}

const PARALLEL_TURN_ERROR = "当前已有会话正在执行，请等待完成或停止后再开始其他会话";
const NEW_THREAD_RUNNING_ERROR = "当前已有会话正在执行，请等待完成或停止后再新建会话";
const agentIpcLogger = createLogger("AgentIpcHandlers");

async function resolveAgentLoop(
  deps: AgentIpcHandlerDeps,
  threadId?: string | null,
): Promise<AgentLoop | null> {
  const manager = deps.agentLoopManagerRef();
  return manager ? await manager.getLoopForThread(threadId) : null;
}

export async function listThreadsForIpc(
  deps: Pick<AgentIpcHandlerDeps, "getSessionStoreInstance" | "getStateRuntimeStoreInstance">,
) {
  if (deps.getStateRuntimeStoreInstance) {
    try {
      const repository = new ThreadRepository(
        deps.getSessionStoreInstance(),
        await deps.getStateRuntimeStoreInstance(),
      );
      return await repository.list();
    } catch (error) {
      agentIpcLogger.warn(
        "SQLite 会话快照不可用，回退读取 JSONL",
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) },
      );
    }
  }
  return deps.getSessionStoreInstance().listThreads();
}

export async function deleteThreadForIpc(
  deps: Pick<AgentIpcHandlerDeps, "getSessionStoreInstance" | "getStateRuntimeStoreInstance">,
  threadId: ThreadId,
): Promise<boolean> {
  if (!deps.getStateRuntimeStoreInstance) {
    return deps.getSessionStoreInstance().deleteThread(threadId);
  }
  const repository = new ThreadRepository(
    deps.getSessionStoreInstance(),
    await deps.getStateRuntimeStoreInstance(),
  );
  return repository.delete(threadId);
}

export async function updateThreadMetadataForIpc(
  deps: Pick<
    AgentIpcHandlerDeps,
    "getSessionStoreInstance" | "getStateRuntimeStoreInstance" | "agentLoopManagerRef"
  >,
  threadId: ThreadId,
  patch: Partial<import("../shared/types").ThreadMetadata>,
) {
  if (!deps.getStateRuntimeStoreInstance) {
    await deps.getSessionStoreInstance().updateThreadMetadata(threadId, patch);
    return;
  }
  const repository = new ThreadRepository(
    deps.getSessionStoreInstance(),
    await deps.getStateRuntimeStoreInstance(),
  );
  const updated = await repository.updateMetadata(threadId, patch);
  deps.agentLoopManagerRef?.()?.updateLoadedThreadMetadata(threadId, updated);
}

export async function enqueueTurnForIpc(
  agent: AgentLoop,
  input: AgentTurnInput,
  callbacks: AgentTurnCallbacks,
  manager?: AgentLoopManager | null,
) {
  if (agent.getIsRunning()) {
    const queued = agent.enqueueTurn(input, callbacks);
    return { success: true, ...queued };
  }

  const targetThreadId = agent.getThread()?.metadata.threadId ?? input.threadId ?? null;
  if (manager?.hasRunningLoopOtherThan?.(targetThreadId)) {
    throw new Error(PARALLEL_TURN_ERROR);
  }

  const run = () => agent.runTurn(input, callbacks);
  const turn = manager?.runWithTurnLock ? await manager.runWithTurnLock(agent, run) : await run();
  const threadId = agent.getThread()?.metadata.threadId;
  return { success: true, queued: false, turnId: turn.turnId, threadId };
}

export async function prepareAgentForStartTurn(
  agent: AgentLoop,
  manager?: AgentLoopManager | null,
): Promise<void> {
  const existingThreadId = agent.getThread()?.metadata.threadId ?? null;
  if (manager?.hasRunningLoopOtherThan?.(existingThreadId)) {
    throw new Error(PARALLEL_TURN_ERROR);
  }
  if (!agent.getThread()) {
    await agent.startThread();
  }
  manager?.rememberLoop(agent);
}

export async function startTurnForIpc(
  agent: AgentLoop,
  input: AgentTurnInput,
  callbacks: AgentTurnCallbacks,
  manager?: AgentLoopManager | null,
) {
  const run = async () => {
    await prepareAgentForStartTurn(agent, manager);
    const turn = await agent.runTurn(input, callbacks);
    manager?.rememberLoop(agent);
    return turn;
  };
  return manager?.runWithTurnLock ? manager.runWithTurnLock(agent, run) : run();
}

export async function prepareNewThreadForIpc(
  deps: Pick<AgentIpcHandlerDeps, "agentLoopManagerRef">,
  folderId?: string,
) {
  const manager = deps.agentLoopManagerRef();
  if (!manager) {
    return { success: false, error: "Agent 未初始化" };
  }
  if (manager.hasRunningLoopOtherThan(null)) {
    return { success: false, error: NEW_THREAD_RUNNING_ERROR };
  }
  manager.prepareNewThread(folderId);
  return { success: true };
}

async function ensureKnowledgeRuntimeForIpc(deps: AgentIpcHandlerDeps): Promise<string | null> {
  if (deps.isDataMigrationInProgress?.()) return "数据存储正在迁移，请稍后重试";
  try {
    const runtime = await deps.ensureKnowledgeRuntime?.();
    return runtime?.error || null;
  } catch (error: any) {
    return error?.message || String(error || "未知错误");
  }
}

function formatKnowledgeUnavailableError(error?: string | null): string {
  return error ? `知识库未初始化：${error}` : "知识库未初始化";
}

export async function reindexAuthorizedKnowledgeSources(
  indexer: Pick<ReturnType<typeof getKnowledgeIndexer> & object, "listSources" | "indexFile">,
  pathAuthorizer: PathAuthorizer,
): Promise<IndexResult[]> {
  const results: IndexResult[] = [];
  for (const source of indexer.listSources()) {
    try {
      const sourcePath = assertAuthorizedPath(pathAuthorizer, source.sourcePath);
      results.push(await indexer.indexFile(sourcePath, { skipUnchanged: false }));
    } catch (error) {
      results.push({
        sourcePath: source.sourcePath,
        success: false,
        error: error instanceof Error ? error.message : String(error),
        entryCount: 0,
        durationMs: 0,
      });
    }
  }
  return results;
}

export function registerAgentIpcHandlers(deps: AgentIpcHandlerDeps): void {
  registerToolApprovalHandlers();

  ipcMain.handle("agent:startTurn", async (_event, request: unknown) => {
    const validated = validateInput(AgentStartTurnInput, request);
    const agent = await resolveAgentLoop(deps, validated.threadId);
    if (!agent) return { success: false, error: "Agent 未初始化" };

    const callbacks = createEventForwarder(deps.mainWindowRef);

    try {
      const manager = deps.agentLoopManagerRef();
      const turn = await startTurnForIpc(agent, validated as AgentTurnInput, callbacks, manager);
      const threadId = agent.getThread()?.metadata.threadId;
      return { success: true, turnId: turn.turnId, threadId };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("agent:continueTurn", async (_event, request: unknown) => {
    const validated = validateInput(AgentContinueTurnInput, request);
    const agent = await resolveAgentLoop(deps, validated.threadId);
    if (!agent) return { success: false, error: "Agent 未初始化" };
    const callbacks = createEventForwarder(deps.mainWindowRef);

    try {
      const manager = deps.agentLoopManagerRef();
      const targetThreadId = agent.getThread()?.metadata.threadId ?? validated.threadId ?? null;
      if (manager?.hasRunningLoopOtherThan?.(targetThreadId)) {
        throw new Error(PARALLEL_TURN_ERROR);
      }
      const resume = () =>
        agent.resumeFromInterruption(validated.content, callbacks, validated.attachments);
      const turn = manager?.runWithTurnLock
        ? await manager.runWithTurnLock(agent, resume)
        : await resume();
      manager?.rememberLoop(agent);
      const threadId = agent.getThread()?.metadata.threadId;
      return { success: true, turnId: turn.turnId, threadId };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("agent:enqueueTurn", async (_event, request: unknown) => {
    const validated = validateInput(AgentStartTurnInput, request);
    const agent = await resolveAgentLoop(deps, validated.threadId);
    if (!agent) return { success: false, error: "Agent 未初始化" };
    const callbacks = createEventForwarder(deps.mainWindowRef);

    try {
      const manager = deps.agentLoopManagerRef();
      const result = await enqueueTurnForIpc(
        agent,
        validated as AgentTurnInput,
        callbacks,
        manager,
      );
      manager?.rememberLoop(agent);
      return result;
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("agent:interrupt", async (event, request?: { threadId?: string | null }) => {
    const validated = validateInput(AgentInterruptInput, request);
    const requestId = `ipc-${event.processId}-${event.frameId}-${Date.now()}`;
    const manager = deps.agentLoopManagerRef();
    if (manager) {
      const interrupted = await manager.interruptThread(validated?.threadId, requestId);
      return interrupted ? { success: true } : { success: false, error: "没有正在运行的 Agent" };
    }
    return { success: false, error: "没有正在运行的 Agent" };
  });

  ipcMain.handle("thread:list", async () => {
    return listThreadsForIpc(deps);
  });

  ipcMain.handle("thread:load", async (_event, threadId: unknown) => {
    const validated = validateInput(ThreadIdInput, threadId);
    return deps.getSessionStoreInstance().loadThread(validated);
  });

  ipcMain.handle("thread:delete", async (_event, threadId: unknown) => {
    const validated = validateInput(ThreadIdInput, threadId);
    const manager = deps.agentLoopManagerRef();
    if (manager && !(await manager.releaseThread(validated))) {
      throw new Error("会话正在运行，停止后才能删除");
    }
    return deleteThreadForIpc(deps, validated);
  });

  ipcMain.handle("thread:resume", async (_event, threadId: unknown) => {
    const validated = validateInput(ThreadIdInput, threadId);
    const manager = deps.agentLoopManagerRef();
    const agent = await resolveAgentLoop(deps, validated);
    if (!agent) return { success: false };
    if (!manager) {
      return { success: await agent.resumeThread(validated) };
    }
    manager.rememberLoop(agent);
    return { success: true };
  });

  ipcMain.handle("thread:new", async (_event, folderId?: unknown) => {
    const validated = validateInput(ThreadNewInput, folderId);
    return prepareNewThreadForIpc(deps, validated);
  });

  ipcMain.handle("thread:updateMetadata", async (_event, threadId: unknown, patch: unknown) => {
    const validated = validateInput(ThreadUpdateMetadataInput, { threadId, patch });
    return updateThreadMetadataForIpc(deps, validated.threadId, validated.patch);
  });

  ipcMain.handle("thread:findLatest", async () => {
    return deps.getSessionStoreInstance().findLatestThread();
  });

  ipcMain.handle("thread:runtimeStatus", async () => {
    const agent = deps.agentLoopManagerRef()?.getPrimaryLoop();
    return agent?.getThreadRuntimeStatus() ?? { status: "not_loaded", idleUnloadMs: 0 };
  });

  ipcMain.handle("threadGraph:upsertSpawnEdge", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphEdgeInput, request);
    return deps
      .getAgentGraphStoreInstance()
      .upsertThreadSpawnEdge(validated.parentThreadId, validated.childThreadId, {
        label: validated.label,
      });
  });

  ipcMain.handle("threadGraph:closeSpawnEdge", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphCloseEdgeInput, request);
    return deps
      .getAgentGraphStoreInstance()
      .closeThreadSpawnEdge(validated.parentThreadId, validated.childThreadId);
  });

  ipcMain.handle("threadGraph:listDescendants", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphListDescendantsInput, request);
    return deps
      .getAgentGraphStoreInstance()
      .listThreadSpawnDescendants(validated.parentThreadId, { status: validated.status });
  });

  ipcMain.handle("tools:list", () => ALL_TOOL_DEFINITIONS);

  ipcMain.handle("stats:getSummary", async (_event, options?: unknown) => {
    validateInput(StatsGetSummaryInput, options);
    return deps.getSessionStoreInstance().getUsageSummary();
  });

  ipcMain.handle("knowledge:listSources", async () => {
    if (deps.isDataMigrationInProgress?.()) throw new Error("数据存储正在迁移，请稍后重试");
    let store = getKnowledgeStore();
    const initError = !store ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    store = getKnowledgeStore();
    if (!store) throw new Error(formatKnowledgeUnavailableError(initError));
    return store.listSources();
  });

  ipcMain.handle("knowledge:search", async (_event, query: unknown, topK: unknown) => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    const validated = validateInput(KnowledgeSearchInput, { query, topK });
    let retriever = getKnowledgeRetriever();
    const initError = !retriever ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    retriever = getKnowledgeRetriever();
    if (!retriever) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await retriever.search({
        text: validated.query,
        topK: validated.topK || 5,
      });
      return { success: true, data: results };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("knowledge:indexFile", async (_event, filePath: unknown) => {
    if (deps.isDataMigrationInProgress?.())
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    const validated = validateInput(KnowledgeIndexFileInput, { filePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const sourcePath = assertAuthorizedPath(deps.pathAuthorizer, validated.filePath);
      const result = await indexer.indexFile(sourcePath);
      return result;
    } catch (err: any) {
      return {
        success: false,
        error: err.message,
        sourcePath: validated.filePath,
        entryCount: 0,
        durationMs: 0,
      };
    }
  });

  ipcMain.handle("knowledge:indexFolder", async (_event, folderPath: unknown) => {
    if (deps.isDataMigrationInProgress?.())
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    const validated = validateInput(KnowledgeIndexFolderInput, { folderPath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const folderPath = assertAuthorizedPath(deps.pathAuthorizer, validated.folderPath);
      const results = await indexer.indexFolder(folderPath);
      return results;
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("knowledge:deleteFile", async (_event, sourcePath: unknown) => {
    if (deps.isDataMigrationInProgress?.())
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    const validated = validateInput(KnowledgeDeleteInput, { sourcePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const sourcePath = assertAuthorizedPath(deps.pathAuthorizer, validated.sourcePath);
      await indexer.deleteSource(sourcePath);
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("knowledge:reindexAll", async () => {
    if (deps.isDataMigrationInProgress?.())
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await reindexAuthorizedKnowledgeSources(indexer, deps.pathAuthorizer);
      return { success: true, results };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });
}
