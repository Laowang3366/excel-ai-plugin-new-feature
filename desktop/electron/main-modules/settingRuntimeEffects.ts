/** 已持久化设置向 Agent、知识库和窗口运行时的同步。 */

import type { BrowserWindow } from "electron";

import type { AgentLoop } from "../agent/core/agentLoop";
import type { AIClientConfig } from "../agent/providers/aiClientTypes";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import {
  buildCompactionConfig,
  type SavedCompactionConfig,
} from "../agent/runtime/compactionRuntime";
import { createLogger } from "../shared/logger";

type RuntimeAgent = Pick<
  AgentLoop,
  "updateAIConfig" | "updateCompactionConfig" | "updatePermissionMode"
>;

export interface SettingRuntimeEffectsDeps {
  agents: RuntimeAgent[];
  mainWindow: BrowserWindow | null;
  getActiveAIConfig: () => AIClientConfig;
  getActiveDataPath: () => string;
  getRuntimeSettingValue: (key: string) => unknown;
  refreshKnowledgeRuntime: (aiConfig: AIClientConfig, dataRoot: string) => Promise<unknown>;
  applyWindowTheme: (mainWindow: BrowserWindow | null) => void;
  applyWindowOpacity: (mainWindow: BrowserWindow | null) => void;
  setDynamicArrayFunctionsEnabled: (value: unknown) => void;
}

const settingRuntimeLogger = createLogger("SettingRuntimeEffects");

export async function applySettingRuntimeEffects(
  key: string,
  value: unknown,
  deps: SettingRuntimeEffectsDeps,
): Promise<void> {
  if (key === "activeProvider" || key === "aiProviders") {
    const aiConfig = deps.getActiveAIConfig();
    const compactionConfig = buildCompactionConfig({
      contextWindowSize: aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
      savedCompaction: deps.getRuntimeSettingValue("compactionConfig") as
        SavedCompactionConfig | undefined,
    });
    for (const agent of deps.agents) {
      agent.updateAIConfig(aiConfig);
      agent.updateCompactionConfig(compactionConfig);
    }
    try {
      await deps.refreshKnowledgeRuntime(aiConfig, deps.getActiveDataPath());
    } catch (error) {
      settingRuntimeLogger.warn("刷新知识库运行时失败，设置已保存:", error);
    }
  }

  if (key === "permissionMode") {
    for (const agent of deps.agents) {
      agent.updatePermissionMode(value as "normal" | "auto_approve_safe" | "confirm_all");
    }
  }

  if (key === "compactionConfig") {
    const aiConfig = deps.getActiveAIConfig();
    const compactionConfig = buildCompactionConfig({
      contextWindowSize: aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
      savedCompaction: deps.getRuntimeSettingValue("compactionConfig") as
        SavedCompactionConfig | undefined,
    });
    for (const agent of deps.agents) agent.updateCompactionConfig(compactionConfig);
  }

  if (key === "theme") deps.applyWindowTheme(deps.mainWindow);
  if (key === "windowOpacity") deps.applyWindowOpacity(deps.mainWindow);
  if (key === "dynamicArrayFunctionsEnabled") deps.setDynamicArrayFunctionsEnabled(value);
}
