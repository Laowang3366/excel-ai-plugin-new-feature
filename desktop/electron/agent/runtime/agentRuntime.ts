import { AgentLoop, type AgentLoopConfig } from "../core/agentLoop";
import type { AIClientConfig } from "../providers/aiClient";
import { DEFAULT_CONTEXT_WINDOW } from "../providers/modelContextWindows";
import { buildSystemPrompt } from "../prompts/systemPrompt";
import { SessionStore } from "../memory/sessionStore";
import type { StateRuntimeStore } from "../memory/stateRuntimeStore";
import { LongTermMemoryStore } from "../memory/longTerm/memoryStore";
import { createToolExecutors } from "../tools/executors/createToolExecutors";
import { OfficeComActionBridge } from "../tools/implementations/office/officeComActionBridge";
import { createOfficeActionBridge } from "../tools/officeCore/officeActionAdapter";
import { getOrCreateOfficeBridges, type OfficeBridgeRegistry } from "./bridgeRegistry";
import { buildCompactionConfig, type SavedCompactionConfig } from "./compactionRuntime";
import { initializeKnowledgeRuntime, reloadKnowledgeRuntime, type KnowledgeRuntimeState } from "./knowledgeRuntime";

export interface AgentRuntime {
  agentLoop: AgentLoop;
  agentLoopManager: AgentLoopManager;
  bridges: OfficeBridgeRegistry;
  knowledge: KnowledgeRuntimeState;
  stateRuntime: StateRuntimeStore;
}

export interface AgentRuntimeDependencies {
  getActiveAIConfig: () => AIClientConfig;
  getActiveDataPath?: () => string;
  getSettingsValue: (key: string) => unknown;
  getSessionStoreInstance: () => SessionStore;
  getStateRuntimeStoreInstance: () => Promise<StateRuntimeStore>;
  requestToolApproval: NonNullable<AgentLoopConfig["requestToolApproval"]>;
}

let runtime: AgentRuntime | null = null;

export class AgentLoopManager {
  private readonly loopsByThreadId = new Map<string, AgentLoop>();
  private pendingFolderId: string | undefined;

  constructor(
    private readonly createLoop: () => AgentLoop,
    private readonly primaryLoop: AgentLoop
  ) {}

  getPrimaryLoop(): AgentLoop {
    return this.primaryLoop;
  }

  getAllLoops(): AgentLoop[] {
    return Array.from(new Set([this.primaryLoop, ...this.loopsByThreadId.values()]));
  }

  hasRunningLoopOtherThan(threadId?: string | null): boolean {
    const targetThreadId = threadId ?? null;
    return this.getAllLoops().some((loop) => {
      if (!loop.getIsRunning()) return false;
      const loopThreadId = loop.getThread()?.metadata.threadId ?? null;
      return !targetThreadId || loopThreadId !== targetThreadId;
    });
  }

  rememberLoop(loop: AgentLoop): void {
    const threadId = loop.getThread()?.metadata.threadId;
    if (threadId) {
      this.loopsByThreadId.set(threadId, loop);
    }
  }

  prepareNewThread(folderId?: string): void {
    this.pendingFolderId = folderId || undefined;
  }

  async getLoopForThread(threadId?: string | null): Promise<AgentLoop> {
    if (!threadId) {
      const loop = this.createLoop();
      await loop.resetThread(this.pendingFolderId);
      this.pendingFolderId = undefined;
      return loop;
    }

    const existing = this.loopsByThreadId.get(threadId);
    if (existing) return existing;

    const loop = this.createLoop();
    const resumed = await loop.resumeThread(threadId);
    if (!resumed) {
      throw new Error("会话不存在或无法恢复");
    }
    this.rememberLoop(loop);
    return loop;
  }

  async interruptThread(threadId?: string | null, requestId?: string): Promise<boolean> {
    if (threadId) {
      const loop = this.loopsByThreadId.get(threadId);
      if (!loop) return false;
      await loop.interrupt(requestId);
      return true;
    }

    const runningLoops = this.getAllLoops().filter((loop) => loop.getIsRunning());
    await Promise.all(runningLoops.map((loop) => loop.interrupt(requestId)));
    return runningLoops.length > 0;
  }
}

/**
 * Agent 运行时总装配。
 *
 * 关联模块：
 * - runtime/bridgeRegistry: 提供 Office/Excel 桥接实例。
 * - runtime/knowledgeRuntime: 初始化 RAG 知识库。
 * - tools/executors/createToolExecutors: 根据桥接和知识检索器组装工具执行器。
 * - core/agentLoop: 消费最终配置并驱动模型自主调用工具。
 */
export async function getOrCreateAgentRuntime(deps: AgentRuntimeDependencies): Promise<AgentRuntime> {
  if (runtime) return runtime;

  const bridges = getOrCreateOfficeBridges();
  const aiConfig = deps.getActiveAIConfig();
  const stateRuntime = await deps.getStateRuntimeStoreInstance();
  const knowledge = await initializeKnowledgeRuntime(aiConfig, deps.getActiveDataPath?.());
  const memoryStore = new LongTermMemoryStore(stateRuntime);
  const officeActionBridge = createOfficeActionBridge({
    officeFileBridge: bridges.officeFileBridge,
    officeComActionBridge: new OfficeComActionBridge(),
  });
  const toolExecutors = createToolExecutors(
    bridges.excelBridge,
    bridges.vbaBridge,
    bridges.scriptBridge,
    bridges.uiBridge,
    undefined,
    knowledge.retriever ?? undefined,
    bridges.wordBridge,
    bridges.presentationBridge,
    bridges.officeScriptBridge,
    officeActionBridge,
    memoryStore,
    {
      getMineruApiToken: () => {
        const configured = deps.getSettingsValue("mineruApiToken") || deps.getSettingsValue("ocrMineruApiToken");
        return typeof configured === "string" ? configured : "";
      },
    }
  );

  const createAgentLoop = () => {
    const activeAiConfig = deps.getActiveAIConfig();
    const activeContextWindowSize = activeAiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW;
    const activeCompactionConfig = buildCompactionConfig({
      contextWindowSize: activeContextWindowSize,
      savedCompaction: deps.getSettingsValue("compactionConfig") as SavedCompactionConfig | undefined,
    });
    return new AgentLoop({
      aiConfig: activeAiConfig,
      systemPrompt: buildSystemPrompt(),
      compactionConfig: activeCompactionConfig,
      toolExecutors,
      memoryStore,
      permissionMode: deps.getSettingsValue("permissionMode") as "normal" | "auto_approve_safe" | "confirm_all" || "normal",
      requestToolApproval: deps.requestToolApproval,
    }, deps.getSessionStoreInstance(), stateRuntime);
  };
  const agentLoop = createAgentLoop();
  const agentLoopManager = new AgentLoopManager(createAgentLoop, agentLoop);

  runtime = { agentLoop, agentLoopManager, bridges, knowledge, stateRuntime };
  return runtime;
}

export function getAgentLoop(): AgentLoop | null {
  return runtime?.agentLoop ?? null;
}

export function getAgentLoops(): AgentLoop[] {
  return runtime?.agentLoopManager.getAllLoops() ?? [];
}

export function getAgentLoopManager(): AgentLoopManager | null {
  return runtime?.agentLoopManager ?? null;
}

export async function refreshKnowledgeRuntime(
  aiConfig: AIClientConfig,
  dataRoot?: string
): Promise<KnowledgeRuntimeState> {
  const knowledge = await reloadKnowledgeRuntime(aiConfig, dataRoot);
  if (runtime) {
    runtime.knowledge = knowledge;
  }
  return knowledge;
}
