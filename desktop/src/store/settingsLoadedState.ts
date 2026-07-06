import type { AiProviderConfig } from "../electronApi";
import type { AppLanguage, AppTheme, PermissionMode, PinnedFolder, SettingsState } from "./settingsStore";
import { normalizeProviderConfig } from "./settingsProviderState";
import { normalizeWindowOpacity } from "./settingsValues";

export interface LoadedSettingsState {
  patch: Partial<SettingsState>;
  migratedProviders?: Record<string, AiProviderConfig>;
}

export function buildLoadedSettingsState(allSettings: Record<string, any>): LoadedSettingsState {
  const patch: Partial<SettingsState> = {};
  let migratedProviders: Record<string, AiProviderConfig> | undefined;

  if (allSettings.aiProviders) {
    const providers = { ...(allSettings.aiProviders as Record<string, AiProviderConfig>) };
    let needsMigration = false;
    for (const id of Object.keys(providers)) {
      let provider = providers[id];
      if (!provider.reasoningMode) {
        needsMigration = true;
        provider = { ...provider, reasoningMode: provider.enableReasoning ? "high" : "off" };
      }
      const normalized = normalizeProviderConfig(provider);
      if (JSON.stringify(normalized) !== JSON.stringify(providers[id])) {
        needsMigration = true;
        providers[id] = normalized;
      }
    }
    patch.providers = providers;
    if (needsMigration) migratedProviders = providers;
  }

  if (allSettings.activeProvider) {
    patch.activeProviderId = allSettings.activeProvider;
  }
  if (allSettings.permissionMode) {
    patch.permissionMode = allSettings.permissionMode as PermissionMode;
  }
  if (typeof allSettings.showReasoning === "boolean") {
    patch.showReasoning = allSettings.showReasoning;
  }
  if (allSettings.language === "zh-CN" || allSettings.language === "en-US") {
    patch.language = allSettings.language as AppLanguage;
  }
  if (allSettings.theme === "light" || allSettings.theme === "dark") {
    patch.theme = allSettings.theme as AppTheme;
  }
  if (typeof allSettings.closeToTray === "boolean") {
    patch.closeToTray = allSettings.closeToTray;
  }
  if (typeof allSettings.officeAutoCompactEnabled === "boolean") {
    patch.officeAutoCompactEnabled = allSettings.officeAutoCompactEnabled;
  }
  if (typeof allSettings.dynamicArrayFunctionsEnabled === "boolean") {
    patch.dynamicArrayFunctionsEnabled = allSettings.dynamicArrayFunctionsEnabled;
  }
  if (allSettings.windowOpacity !== undefined) {
    patch.windowOpacity = normalizeWindowOpacity(allSettings.windowOpacity);
  }

  const compactionConfig = allSettings.compactionConfig;
  if (compactionConfig && typeof compactionConfig === "object") {
    if (typeof compactionConfig.enabled === "boolean") {
      patch.compactionEnabled = compactionConfig.enabled;
    }
    if (typeof compactionConfig.autoCompactThresholdPercent === "number") {
      patch.autoCompactThresholdPercent = compactionConfig.autoCompactThresholdPercent;
    }
  }

  if (Array.isArray(allSettings.pinnedFolders)) {
    patch.pinnedFolders = allSettings.pinnedFolders as PinnedFolder[];
  }
  if (typeof allSettings.knowledgeEnabled === "boolean") {
    patch.knowledgeEnabled = allSettings.knowledgeEnabled;
  }

  return { patch, migratedProviders };
}
