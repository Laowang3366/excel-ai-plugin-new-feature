import type { AgentLoop } from "../agent/core/agentLoop";
import type { AIClientConfig } from "../agent/providers/aiClient";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import {
  buildCompactionConfig,
  type SavedCompactionConfig,
} from "../agent/runtime/compactionRuntime";
import type { SessionStore } from "../agent/memory/sessionStore";
import { getActiveDataPath } from "./settingsDataPath";
import {
  afterUserDataExported,
  rotateLocalDataKey,
} from "./localDataProtection/localDataLifecycle";
import { eraseAllManagedReplicasAndKeys } from "./localDataProtection/localDataEraseAll";
import { runUserDataExport } from "./userDataExportCoordinator";
import { runUserDataErase } from "./userDataEraseCoordinator";

/** Narrow deps injected by settingsManager — no reverse import of settingsManager. */
export interface SettingsLocalDataContext {
  isBusy: () => boolean;
  setBusy: (busy: boolean) => void;
  getAgents: () => AgentLoop[];
  getSessionStore: () => SessionStore;
  closeStateRuntime: () => Promise<void>;
  resetKnowledge: () => void;
  clearSettings: () => void;
  resetSessionStore: () => Promise<void>;
  getActiveAIConfig: () => AIClientConfig;
  getRuntimeSettingValue: (key: string) => unknown;
  reloadKnowledge: (
    config: AIClientConfig,
    dataPath: string,
  ) => Promise<{ store: unknown; error?: string | null }>;
  getSanitizedSettings: () => Record<string, unknown>;
}

export async function runExportUserData(
  targetPath: string,
  ctx: SettingsLocalDataContext,
): ReturnType<typeof runUserDataExport> {
  const result = await runUserDataExport(targetPath, {
    isBusy: ctx.isBusy,
    setBusy: ctx.setBusy,
    hasRunningAgent: () => ctx.getAgents().some((agent) => agent.getIsRunning()),
    getDataPath: getActiveDataPath,
    getSanitizedSettings: ctx.getSanitizedSettings,
    getSessionStore: ctx.getSessionStore,
    closeStateRuntime: ctx.closeStateRuntime,
    resetKnowledgeRuntime: ctx.resetKnowledge,
    restoreRuntimes: async () => {
      await ctx.resetSessionStore();
      await ctx.reloadKnowledge(ctx.getActiveAIConfig(), getActiveDataPath());
    },
  });
  // Register any committed export path even when runtime restore failed (success:false + exportPath).
  if (result.exportPath) {
    afterUserDataExported(result.exportPath);
  }
  return result;
}

export async function runEraseUserData(
  confirmation: string,
  ctx: SettingsLocalDataContext,
): Promise<{
  success: boolean;
  erasedCategories: string[];
  errors: string[];
  error?: string;
  proofSummary?: {
    createdAt: string;
    installIdDigest: string;
    proofDigest: string;
    destroyedKeyCount: number;
    keyMaterialDestroyed: boolean;
    replicaCount: number;
    erasedCount: number;
    failedCount: number;
  };
}> {
  return runUserDataErase(confirmation, {
    isBusy: ctx.isBusy,
    setBusy: ctx.setBusy,
    hasRunningAgent: () => ctx.getAgents().some((agent) => agent.getIsRunning()),
    getDataPath: getActiveDataPath,
    getSessionStore: ctx.getSessionStore,
    resetAgents: async () => {
      await Promise.all(ctx.getAgents().map((agent) => agent.resetThread()));
    },
    closeStateRuntime: ctx.closeStateRuntime,
    resetKnowledgeRuntime: ctx.resetKnowledge,
    clearSettings: ctx.clearSettings,
    // Single ownership: eraseAll covers active/old/export once; coordinator only clears settings.
    eraseManagedData: async () => ({ erasedCategories: [], errors: [] }),
    // Key swap is inside eraseAll (pending→commit→purge); no null-protection reseed window.
    afterQuiesceBeforeRestore: async () => {
      const replicaErase = await eraseAllManagedReplicasAndKeys();
      return {
        erasedCategories: replicaErase.erasedCategories,
        errors: replicaErase.errors,
        proofSummary: replicaErase.proofSummary,
      };
    },
    restoreRuntimes: async () => {
      await ctx.resetSessionStore();
      const aiConfig = ctx.getActiveAIConfig();
      const compactionConfig = buildCompactionConfig({
        contextWindowSize: aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
        savedCompaction: ctx.getRuntimeSettingValue("compactionConfig") as
          SavedCompactionConfig | undefined,
      });
      for (const agent of ctx.getAgents()) {
        agent.updateAIConfig(aiConfig);
        agent.updateCompactionConfig(compactionConfig);
        agent.updatePermissionMode("normal");
      }
      const restoredKnowledge = await ctx.reloadKnowledge(aiConfig, getActiveDataPath());
      if (!restoredKnowledge.store) {
        throw new Error(restoredKnowledge.error || "知识库恢复失败");
      }
    },
  }) as Promise<{
    success: boolean;
    erasedCategories: string[];
    errors: string[];
    error?: string;
    proofSummary?: {
      createdAt: string;
      installIdDigest: string;
      proofDigest: string;
      destroyedKeyCount: number;
      keyMaterialDestroyed: boolean;
      replicaCount: number;
      erasedCount: number;
      failedCount: number;
    };
  }>;
}

export async function runRotateLocalDataEncryptionKey(
  ctx: SettingsLocalDataContext,
): Promise<{ success: boolean; keyId?: number; error?: string }> {
  if (ctx.isBusy()) {
    return { success: false, error: "数据维护或相关操作正在进行中，请稍后重试" };
  }
  if (ctx.getAgents().some((agent) => agent.getIsRunning())) {
    return { success: false, error: "请等待当前会话执行完成或停止后再轮换密钥" };
  }
  ctx.setBusy(true);
  try {
    return await rotateLocalDataKey({
      dataRoot: getActiveDataPath(),
      seal: async () => {
        ctx.getSessionStore().suspendWrites("正在轮换本地数据密钥");
        await ctx.getSessionStore().flushRolloutWrites();
        await ctx.closeStateRuntime();
        ctx.resetKnowledge();
      },
      restore: async () => {
        ctx.getSessionStore().resumeWrites();
        await ctx.resetSessionStore();
        await ctx.reloadKnowledge(ctx.getActiveAIConfig(), getActiveDataPath());
      },
    });
  } finally {
    ctx.setBusy(false);
  }
}
