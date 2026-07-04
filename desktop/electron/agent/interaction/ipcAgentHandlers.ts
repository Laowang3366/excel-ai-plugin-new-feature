/**
 * Agent 相关 IPC 注册。
 *
 * 关联模块：
 * - interaction/eventForwarder: Agent 事件和工具审批 IPC。
 * - core/agentLoop: 启动、继续、中断、恢复线程。
 * - memory/sessionStore: 线程列表、元数据和统计。
 * - knowledge/knowledgeRegistry: RAG 知识库查询与索引入口。
 */

import { BrowserWindow, ipcMain } from "electron";
import type { AgentLoop } from "../core/agentLoop";
import type { AgentLoopManager } from "../runtime/agentRuntime";
import { getKnowledgeIndexer, getKnowledgeRetriever, getKnowledgeStore } from "../knowledge/knowledgeRegistry";
import type { KnowledgeRuntimeState } from "../runtime/knowledgeRuntime";
import type { SessionStore } from "../memory/sessionStore";
import type { AgentGraphStore } from "../memory/agentGraphStore";
import type { StateRuntimeStore } from "../memory/stateRuntimeStore";
import type { AgentTurnCallbacks, AgentTurnInput, ThreadId } from "../shared/types";
import { ALL_TOOL_DEFINITIONS } from "../tools/registry/toolDefinitions";
import {
  AgentContinueTurnInput,
  AgentStartTurnInput,
  KnowledgeDeleteInput,
  KnowledgeIndexFileInput,
  KnowledgeIndexFolderInput,
  KnowledgeSearchInput,
  ThreadGraphCloseEdgeInput,
  ThreadGraphEdgeInput,
  ThreadGraphListDescendantsInput,
  ThreadUpdateMetadataInput,
  validateInput,
} from "../../shared/ipcSchemas";
import { createEventForwarder, registerToolApprovalHandlers } from "./eventForwarder";

export interface AgentIpcHandlerDeps {
  mainWindowRef: () => BrowserWindow | null;
  agentLoopRef: () => AgentLoop | null;
  agentLoopManagerRef?: () => AgentLoopManager | null;
  getSessionStoreInstance: () => SessionStore;
  getStateRuntimeStoreInstance?: () => Promise<StateRuntimeStore>;
  getAgentGraphStoreInstance: () => AgentGraphStore;
  ensureKnowledgeRuntime?: () => Promise<KnowledgeRuntimeState>;
}

const PARALLEL_TURN_ERROR = "当前已有会话正在执行，请等待完成或停止后再开始其他会话";
const NEW_THREAD_RUNNING_ERROR = "当前已有会话正在执行，请等待完成或停止后再新建会话";

async function resolveAgentLoop(
  deps: AgentIpcHandlerDeps,
  threadId?: string | null
): Promise<AgentLoop | null> {
  const manager = deps.agentLoopManagerRef?.();
  if (manager) {
    return await manager.getLoopForThread(threadId);
  }
  return deps.agentLoopRef();
}

export async function listThreadsForIpc(
  deps: Pick<AgentIpcHandlerDeps, "getSessionStoreInstance" | "getStateRuntimeStoreInstance">
) {
  if (deps.getStateRuntimeStoreInstance) {
    try {
      return await (await deps.getStateRuntimeStoreInstance()).listThreadSnapshots();
    } catch {
      // SQLite 快照不可用时保留 JSONL 兼容路径。
    }
  }
  return deps.getSessionStoreInstance().listThreads();
}

export async function enqueueTurnForIpc(
  agent: AgentLoop,
  input: AgentTurnInput,
  callbacks: AgentTurnCallbacks,
  manager?: AgentLoopManager | null
) {
  if (agent.getIsRunning()) {
    const queued = agent.enqueueTurn(input, callbacks);
    return { success: true, ...queued };
  }

  const targetThreadId = agent.getThread()?.metadata.threadId ?? input.threadId ?? null;
  if (manager?.hasRunningLoopOtherThan?.(targetThreadId)) {
    throw new Error(PARALLEL_TURN_ERROR);
  }

  const turn = await agent.runTurn(input, callbacks);
  const threadId = agent.getThread()?.metadata.threadId;
  return { success: true, queued: false, turnId: turn.turnId, threadId };
}

export async function prepareAgentForStartTurn(
  agent: AgentLoop,
  manager?: AgentLoopManager | null
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

export async function prepareNewThreadForIpc(deps: Pick<AgentIpcHandlerDeps, "agentLoopManagerRef" | "agentLoopRef">, folderId?: string) {
  const manager = deps.agentLoopManagerRef?.();
  if (manager) {
    if (manager.hasRunningLoopOtherThan(null)) {
      return { success: false, error: NEW_THREAD_RUNNING_ERROR };
    }
    manager.prepareNewThread(folderId);
    return { success: true };
  }

  const agent = deps.agentLoopRef();
  if (agent?.getIsRunning()) {
    return { success: false, error: NEW_THREAD_RUNNING_ERROR };
  }
  if (agent) {
    await agent.resetThread(folderId);
  }
  return { success: true };
}

async function ensureKnowledgeRuntimeForIpc(deps: AgentIpcHandlerDeps): Promise<string | null> {
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

export function registerAgentIpcHandlers(deps: AgentIpcHandlerDeps): void {
  registerToolApprovalHandlers();

  ipcMain.handle("agent:startTurn", async (_event, request: unknown) => {
    const validated = validateInput(AgentStartTurnInput, request);
    const agent = await resolveAgentLoop(deps, validated.threadId);
    if (!agent) return { success: false, error: "Agent 未初始化" };

    const callbacks = createEventForwarder(deps.mainWindowRef);

    try {
      await prepareAgentForStartTurn(agent, deps.agentLoopManagerRef?.());
      const turn = await agent.runTurn(validated as AgentTurnInput, callbacks);
      deps.agentLoopManagerRef?.()?.rememberLoop(agent);
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
      const manager = deps.agentLoopManagerRef?.();
      const targetThreadId = agent.getThread()?.metadata.threadId ?? validated.threadId ?? null;
      if (manager?.hasRunningLoopOtherThan?.(targetThreadId)) {
        throw new Error(PARALLEL_TURN_ERROR);
      }
      const turn = await agent.resumeFromInterruption(validated.content, callbacks, validated.attachments);
      deps.agentLoopManagerRef?.()?.rememberLoop(agent);
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
      const result = await enqueueTurnForIpc(agent, validated as AgentTurnInput, callbacks, deps.agentLoopManagerRef?.());
      deps.agentLoopManagerRef?.()?.rememberLoop(agent);
      return result;
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("agent:interrupt", async (event, request?: { threadId?: string | null }) => {
    const requestId = `ipc-${event.processId}-${event.frameId}-${Date.now()}`;
    const manager = deps.agentLoopManagerRef?.();
    if (manager) {
      const interrupted = await manager.interruptThread(request?.threadId, requestId);
      return interrupted
        ? { success: true }
        : { success: false, error: "没有正在运行的 Agent" };
    }
    const agent = deps.agentLoopRef();
    if (agent) {
      await agent.interrupt(requestId);
      return { success: true };
    }
    return { success: false, error: "没有正在运行的 Agent" };
  });

  ipcMain.handle("thread:list", async () => {
    return listThreadsForIpc(deps);
  });

  ipcMain.handle("thread:load", async (_event, threadId: ThreadId) => {
    return deps.getSessionStoreInstance().loadThread(threadId);
  });

  ipcMain.handle("thread:delete", async (_event, threadId: ThreadId) => {
    return deps.getSessionStoreInstance().deleteThread(threadId);
  });

  ipcMain.handle("thread:resume", async (_event, threadId: ThreadId) => {
    const agent = await resolveAgentLoop(deps, threadId);
    if (!agent) return { success: false };
    const success = await agent.resumeThread(threadId);
    deps.agentLoopManagerRef?.()?.rememberLoop(agent);
    return { success };
  });

  ipcMain.handle("thread:new", async (_event, folderId?: string) => {
    return prepareNewThreadForIpc(deps, folderId);
  });

  ipcMain.handle("thread:updateMetadata", async (_event, threadId: unknown, patch: unknown) => {
    const validated = validateInput(ThreadUpdateMetadataInput, { threadId, patch });
    return deps.getSessionStoreInstance().updateThreadMetadata(validated.threadId, validated.patch);
  });

  ipcMain.handle("thread:findLatest", async () => {
    return deps.getSessionStoreInstance().findLatestThread();
  });

  ipcMain.handle("thread:runtimeStatus", async () => {
    const agent = deps.agentLoopRef();
    return agent?.getThreadRuntimeStatus() ?? { status: "not_loaded", idleUnloadMs: 0 };
  });

  ipcMain.handle("threadGraph:upsertSpawnEdge", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphEdgeInput, request);
    return deps.getAgentGraphStoreInstance().upsertThreadSpawnEdge(
      validated.parentThreadId,
      validated.childThreadId,
      { label: validated.label }
    );
  });

  ipcMain.handle("threadGraph:closeSpawnEdge", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphCloseEdgeInput, request);
    return deps.getAgentGraphStoreInstance().closeThreadSpawnEdge(
      validated.parentThreadId,
      validated.childThreadId
    );
  });

  ipcMain.handle("threadGraph:listDescendants", async (_event, request: unknown) => {
    const validated = validateInput(ThreadGraphListDescendantsInput, request);
    return deps.getAgentGraphStoreInstance().listThreadSpawnDescendants(
      validated.parentThreadId,
      { status: validated.status }
    );
  });

  ipcMain.handle("tools:list", () => ALL_TOOL_DEFINITIONS);

  ipcMain.handle("stats:getSummary", async () => {
    return deps.getSessionStoreInstance().getUsageSummary();
  });

  ipcMain.handle("knowledge:listSources", async () => {
    let store = getKnowledgeStore();
    const initError = !store ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    store = getKnowledgeStore();
    if (!store) throw new Error(formatKnowledgeUnavailableError(initError));
    return store.listSources();
  });

  ipcMain.handle("knowledge:search", async (_event, query: unknown, topK: unknown) => {
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
    const validated = validateInput(KnowledgeIndexFileInput, { filePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const result = await indexer.indexFile(validated.filePath);
      return result;
    } catch (err: any) {
      return { success: false, error: err.message, sourcePath: validated.filePath, entryCount: 0, durationMs: 0 };
    }
  });

  ipcMain.handle("knowledge:indexFolder", async (_event, folderPath: unknown) => {
    const validated = validateInput(KnowledgeIndexFolderInput, { folderPath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await indexer.indexFolder(validated.folderPath);
      return results;
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("knowledge:deleteFile", async (_event, sourcePath: unknown) => {
    const validated = validateInput(KnowledgeDeleteInput, { sourcePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      await indexer.deleteSource(validated.sourcePath);
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("knowledge:reindexAll", async () => {
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await indexer.reindexAll();
      return { success: true, results };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });
}
