/**
 * 设置状态管理
 *
 * 管理 AI 提供商配置、权限模式、推理显示等设置。
 * 通过 electron-store 持久化到本地。
 *
 * 设计原则：
 * - PROVIDER_TEMPLATES 仅作为"添加供应商"时的预设模板
 * - providers 初始为空，用户自行添加供应商
 * - 每个供应商有唯一 ID（uuid 风格）
 */

import { create } from "zustand";
import type { AiProviderConfig } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import { buildLoadedSettingsState } from "./settingsLoadedState";
import { savePartial } from "./settingsPersistence";
import {
  buildAddedProviderState,
  buildProviderModelsState,
  buildRemovedProviderState,
  buildUpdatedProviderState,
  checkProviderConfigured,
  generateProviderId,
} from "./settingsProviderState";
import { MAX_WINDOW_OPACITY, normalizeWindowOpacity } from "./settingsValues";

export { API_FORMATS, PROVIDER_TEMPLATES } from "./settingsProviderTemplates";
export type { ProviderCategory, ProviderTemplate, ReasoningOption } from "./settingsProviderTemplates";
export { MAX_WINDOW_OPACITY, MIN_WINDOW_OPACITY, normalizeWindowOpacity } from "./settingsValues";

/** 固定的文件夹 */
export interface PinnedFolder {
  path: string;
  name: string;
  addedAt: number;
  /** 置顶的文件路径列表（按顺序） */
  pinnedFiles?: string[];
}

// ============================================================
// 状态类型
// ============================================================

export type PermissionMode = "normal" | "auto_approve_safe" | "confirm_all";
export type AppLanguage = "zh-CN" | "en-US";
export type AppTheme = "light" | "dark";

export interface SettingsState {
  /** 所有已配置的提供商（用户自行添加） */
  providers: Record<string, AiProviderConfig>;
  /** 当前活跃的提供商 ID */
  activeProviderId: string;
  /** 权限模式 */
  permissionMode: PermissionMode;
  /** 是否显示推理过程 */
  showReasoning: boolean;
  /** 界面语言 */
  language: AppLanguage;
  /** 主题 */
  theme: AppTheme;
  /** 关闭窗口时隐藏到托盘 */
  closeToTray: boolean;
  /** Office 操作时是否自动避让为紧凑栏 */
  officeAutoCompactEnabled: boolean;
  /** 是否默认允许公式助手使用动态数组函数 */
  dynamicArrayFunctionsEnabled: boolean;
  /** 是否允许把文件、文本或查询发送到第三方服务 */
  remoteDataProcessingEnabled: boolean;
  /** Main window opacity, from 0.55 to 1. */
  windowOpacity: number;
  /** 是否启用上下文自动压缩 */
  compactionEnabled: boolean;
  /** 自动压缩触发阈值（百分比，如 80 表示 80%） */
  autoCompactThresholdPercent: number;
  /** 是否已完成初始配置 */
  isConfigured: boolean;
  /** 是否正在加载 */
  isLoading: boolean;
  /** 侧边栏固定的文件夹列表 */
  pinnedFolders: PinnedFolder[];
  /** 是否启用知识库 (RAG) */
  knowledgeEnabled: boolean;
}

export interface SettingsActions {
  /** 从 electron-store 加载设置 */
  loadSettings: () => Promise<void>;
  /** 设置活跃提供商 */
  setActiveProvider: (id: string) => void;
  /** 更新提供商配置 */
  updateProvider: (id: string, patch: Partial<AiProviderConfig>) => void;
  /** 添加提供商 */
  addProvider: (config: AiProviderConfig) => void;
  /** 删除提供商 */
  removeProvider: (id: string) => void;
  /** 更新提供商模型列表 */
  setProviderModels: (id: string, models: string[]) => void;
  /** 生成唯一 ID */
  generateId: () => string;
  /** 设置权限模式 */
  setPermissionMode: (mode: PermissionMode) => void;
  /** 切换推理显示 */
  toggleShowReasoning: () => void;
  /** 设置界面语言 */
  setLanguage: (language: AppLanguage) => void;
  /** 设置主题 */
  setTheme: (theme: AppTheme) => void;
  /** 设置关闭窗口行为 */
  setCloseToTray: (enabled: boolean) => void;
  /** 设置 Office 操作时自动避让 */
  setOfficeAutoCompactEnabled: (enabled: boolean) => void;
  /** 设置动态数组函数环境支持 */
  setDynamicArrayFunctionsEnabled: (enabled: boolean) => void;
  /** 设置是否允许远程数据处理 */
  setRemoteDataProcessingEnabled: (enabled: boolean) => void;
  /** Set main window opacity. */
  setWindowOpacity: (opacity: number) => void;
  /** 设置是否启用上下文自动压缩 */
  setCompactionEnabled: (enabled: boolean) => void;
  /** 设置自动压缩触发阈值百分比 */
  setAutoCompactThresholdPercent: (percent: number) => void;
  /** 检查是否已配置 */
  checkConfigured: () => boolean;
  /** 添加固定文件夹 */
  addPinnedFolder: (folder: PinnedFolder) => void;
  /** 删除固定文件夹 */
  removePinnedFolder: (folderPath: string) => void;
  /** 更新固定文件夹（如 pinnedFiles） */
  updatePinnedFolder: (folderPath: string, patch: Partial<PinnedFolder>) => void;
  /** 设置知识库是否启用 */
  setKnowledgeEnabled: (enabled: boolean) => void;
}

// ============================================================
// Zustand Store
// ============================================================

export const useSettingsStore = create<SettingsState & SettingsActions>((set, get) => ({
  // ---- 初始状态 ----
  providers: {},
  activeProviderId: "",
  permissionMode: "normal",
  showReasoning: true,
  language: "zh-CN",
  theme: "light",
  closeToTray: false,
  officeAutoCompactEnabled: false,
  dynamicArrayFunctionsEnabled: true,
  remoteDataProcessingEnabled: false,
  windowOpacity: MAX_WINDOW_OPACITY,
  compactionEnabled: true,
  autoCompactThresholdPercent: 80,
  isConfigured: false,
  isLoading: true,
  pinnedFolders: [],
  knowledgeEnabled: true,

  // ---- Actions ----

  loadSettings: async () => {
    try {
      const allSettings = await ipcApi.settings.getAll() as Record<string, any>;
      const loaded = buildLoadedSettingsState(allSettings);

      set({
        ...loaded.patch,
        isLoading: false,
        isConfigured: checkProviderConfigured(
          loaded.patch.providers ?? get().providers,
          loaded.patch.activeProviderId ?? get().activeProviderId
        ),
      });
      if (loaded.migratedProviders) {
        await ipcApi.settings.set("aiProviders", loaded.migratedProviders);
      }
    } catch {
      set({ isLoading: false });
    }
  },

  setActiveProvider: (id: string) => {
    set({ activeProviderId: id, isConfigured: get().checkConfigured() });
    savePartial(["activeProviderId"], get);
  },

  updateProvider: (id: string, patch: Partial<AiProviderConfig>) => {
    set((s) => buildUpdatedProviderState(s, id, patch));
    void savePartial(["providers"], get).then((persisted) => {
      if (persisted.aiProviders) {
        set({ providers: persisted.aiProviders as Record<string, AiProviderConfig> });
      }
    });
  },

  addProvider: (config: AiProviderConfig) => {
    set((s) => buildAddedProviderState(s, config));
    void savePartial(["providers", "activeProviderId"], get).then((persisted) => {
      if (persisted.aiProviders) {
        set({ providers: persisted.aiProviders as Record<string, AiProviderConfig> });
      }
    });
  },

  removeProvider: (id: string) => {
    set((s) => buildRemovedProviderState(s, id));
    void savePartial(["providers", "activeProviderId"], get).then((persisted) => {
      if (persisted.aiProviders) {
        set({ providers: persisted.aiProviders as Record<string, AiProviderConfig> });
      }
    });
  },

  setProviderModels: (id: string, models: string[]) => {
    set((s) => ({
      providers: buildProviderModelsState(s.providers, id, models),
    }));
    void savePartial(["providers"], get).then((persisted) => {
      if (persisted.aiProviders) {
        set({ providers: persisted.aiProviders as Record<string, AiProviderConfig> });
      }
    });
  },

  generateId: () => {
    return generateProviderId();
  },

  setPermissionMode: (mode: PermissionMode) => {
    set({ permissionMode: mode });
    savePartial(["permissionMode"], get);
  },

  toggleShowReasoning: () => {
    set((s) => ({ showReasoning: !s.showReasoning }));
    savePartial(["showReasoning"], get);
  },

  setLanguage: (language: AppLanguage) => {
    set({ language });
    savePartial(["language"], get);
  },

  setTheme: (theme: AppTheme) => {
    set({ theme });
    savePartial(["theme"], get);
  },

  setCloseToTray: (enabled: boolean) => {
    set({ closeToTray: enabled });
    savePartial(["closeToTray"], get);
  },

  setOfficeAutoCompactEnabled: (enabled: boolean) => {
    set({ officeAutoCompactEnabled: enabled });
    savePartial(["officeAutoCompactEnabled"], get);
  },

  setDynamicArrayFunctionsEnabled: (enabled: boolean) => {
    set({ dynamicArrayFunctionsEnabled: enabled });
    savePartial(["dynamicArrayFunctionsEnabled"], get);
  },

  setRemoteDataProcessingEnabled: (enabled: boolean) => {
    set({ remoteDataProcessingEnabled: enabled });
    savePartial(["remoteDataProcessingEnabled"], get);
  },

  setWindowOpacity: (opacity: number) => {
    set({ windowOpacity: normalizeWindowOpacity(opacity) });
    savePartial(["windowOpacity"], get);
  },

  setCompactionEnabled: (enabled: boolean) => {
    set({ compactionEnabled: enabled });
    savePartial(["compactionEnabled"], get);
  },

  setAutoCompactThresholdPercent: (percent: number) => {
    set({ autoCompactThresholdPercent: percent });
    savePartial(["autoCompactThresholdPercent"], get);
  },

  checkConfigured: () => {
    const { providers, activeProviderId } = get();
    return checkProviderConfigured(providers, activeProviderId);
  },

  addPinnedFolder: (folder: PinnedFolder) => {
    const { pinnedFolders } = get();
    // 避免重复
    if (pinnedFolders.some((f) => f.path === folder.path)) return;
    set({ pinnedFolders: [...pinnedFolders, folder] });
    savePartial(["pinnedFolders"], get);
  },

  removePinnedFolder: (folderPath: string) => {
    const { pinnedFolders } = get();
    set({ pinnedFolders: pinnedFolders.filter((f) => f.path !== folderPath) });
    savePartial(["pinnedFolders"], get);
  },

  updatePinnedFolder: (folderPath: string, patch: Partial<PinnedFolder>) => {
    const { pinnedFolders } = get();
    set({
      pinnedFolders: pinnedFolders.map((f) =>
        f.path === folderPath ? { ...f, ...patch } : f
      ),
    });
    savePartial(["pinnedFolders"], get);
  },

  setKnowledgeEnabled: (enabled: boolean) => {
    set({ knowledgeEnabled: enabled });
    savePartial(["knowledgeEnabled"], get);
  },
}));
