/** 设置 IPC，以及设置写入后的主进程运行时同步。 */

import type { BrowserWindow } from "electron";

import type { AgentLoop } from "../agent/core/agentLoop";
import { setDynamicArrayFunctionsEnabled } from "../agent/runtime/agentGlobalSettings";
import { refreshKnowledgeRuntime } from "../agent/runtime/agentRuntime";
import { SettingsGetInput, SettingsSetInput, validateInput } from "../shared/ipcSchemas";
import { trustedIpcMain as ipcMain } from "../shared/trustedIpc";
import { guardDataOperation } from "./dataMaintenance";
import {
  applyWindowOpacity,
  applyWindowTheme,
  getActiveAIConfig,
  getActiveDataPath,
  getRuntimeSettingValue,
  getSettingForRenderer,
  getSettingsForRenderer,
  isDataMigrationInProgress,
  setSettingFromRenderer,
} from "./settingsManager";
import { applySettingRuntimeEffects } from "./settingRuntimeEffects";

export interface SettingsIpcHandlerDeps {
  mainWindowRef: () => BrowserWindow | null;
  agentLoopsRef: () => AgentLoop[];
}

export function registerSettingsIpcHandlers(deps: SettingsIpcHandlerDeps): void {
  ipcMain.handle("settings:get", (_event, key: unknown) => {
    return getSettingForRenderer(validateInput(SettingsGetInput, key));
  });

  ipcMain.handle(
    "settings:set",
    guardDataOperation(
      isDataMigrationInProgress,
      async (_event, keyInput: unknown, valueInput: unknown) => {
        const [key, value] = validateInput(SettingsSetInput, [keyInput, valueInput]);
        const rendererValue = setSettingFromRenderer(key, value);
        await applySettingRuntimeEffects(key, value, {
          agents: deps.agentLoopsRef(),
          mainWindow: deps.mainWindowRef(),
          getActiveAIConfig,
          getActiveDataPath,
          getRuntimeSettingValue,
          refreshKnowledgeRuntime,
          applyWindowTheme,
          applyWindowOpacity,
          setDynamicArrayFunctionsEnabled,
        });
        return rendererValue;
      },
    ),
  );

  ipcMain.handle("settings:getAll", () => getSettingsForRenderer());
}
