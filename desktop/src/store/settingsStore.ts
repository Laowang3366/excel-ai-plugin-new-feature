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
import { normalizeProviderReasoningConfig } from "../utils/reasoningSupport";
import { savePartial } from "./settingsPersistence";
import { PROVIDER_TEMPLATES, type ProviderTemplate } from "./settingsProviderTemplates";

export { API_FORMATS, PROVIDER_TEMPLATES } from "./settingsProviderTemplates";
export type { ProviderCategory, ProviderTemplate, ReasoningOption } from "./settingsProviderTemplates";

/** 固定的文件夹 */
export interface PinnedFolder {
  path: string;
  name: string;
  addedAt: number;
  /** 置顶的文件路径列表（按顺序） */
  pinnedFiles?: string[];
}

function getProviderTemplate(provider: Pick<AiProviderConfig, "provider">): ProviderTemplate | undefined {
  return PROVIDER_TEMPLATES.find((template) => template.provider === provider.provider);
}

function normalizeProviderConfig(provider: AiProviderConfig): AiProviderConfig {
  return normalizeProviderReasoningConfig(provider, getProviderTemplate(provider));
}

// ============================================================
// 状态类型
// ============================================================

export type PermissionMode = "normal" | "auto_approve_safe" | "confirm_all";
export type AppLanguage = "zh-CN" | "en-US";
export type AppTheme = "light" | "dark";

export const MIN_WINDOW_OPACITY = 0.55;
export const MAX_WINDOW_OPACITY = 1;

export function normalizeWindowOpacity(value: unknown): number {
  const numericValue = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numericValue)) return MAX_WINDOW_OPACITY;
  const clamped = Math.min(Math.max(numericValue, MIN_WINDOW_OPACITY), MAX_WINDOW_OPACITY);
  return Math.round(clamped * 100) / 100;
}

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
  /** 保存设置到 electron-store */
  saveSettings: () => Promise<void>;
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

      if (allSettings.aiProviders) {
        // 数据迁移：enableReasoning → reasoningMode
        const providers = allSettings.aiProviders as Record<string, AiProviderConfig>;
        let needsMigration = false;
        for (const id of Object.keys(providers)) {
          let p = providers[id];
          if (!p.reasoningMode) {
            needsMigration = true;
            p = { ...p, reasoningMode: p.enableReasoning ? "high" : "off" };
          }
          const normalized = normalizeProviderConfig(p);
          if (JSON.stringify(normalized) !== JSON.stringify(providers[id])) {
            needsMigration = true;
            providers[id] = normalized;
          }
        }
        set({ providers });
        // 如果有迁移，立即保存
        if (needsMigration) {
          await ipcApi.settings.set("aiProviders", providers);
        }
      }
      if (allSettings.activeProvider) {
        set({ activeProviderId: allSettings.activeProvider });
      }
      if (allSettings.permissionMode) {
        set({ permissionMode: allSettings.permissionMode });
      }
      if (typeof allSettings.showReasoning === "boolean") {
        set({ showReasoning: allSettings.showReasoning });
      }
      if (allSettings.language === "zh-CN" || allSettings.language === "en-US") {
        set({ language: allSettings.language });
      }
      if (allSettings.theme === "light" || allSettings.theme === "dark") {
        set({ theme: allSettings.theme });
      }
      if (typeof allSettings.closeToTray === "boolean") {
        set({ closeToTray: allSettings.closeToTray });
      }
      if (typeof allSettings.officeAutoCompactEnabled === "boolean") {
        set({ officeAutoCompactEnabled: allSettings.officeAutoCompactEnabled });
      }
      if (typeof allSettings.dynamicArrayFunctionsEnabled === "boolean") {
        set({ dynamicArrayFunctionsEnabled: allSettings.dynamicArrayFunctionsEnabled });
      }
      if (allSettings.windowOpacity !== undefined) {
        set({ windowOpacity: normalizeWindowOpacity(allSettings.windowOpacity) });
      }
      // 上下文压缩配置
      const compactionConfig = allSettings.compactionConfig;
      if (compactionConfig && typeof compactionConfig === "object") {
        if (typeof compactionConfig.enabled === "boolean") {
          set({ compactionEnabled: compactionConfig.enabled });
        }
        if (typeof compactionConfig.autoCompactThresholdPercent === "number") {
          set({ autoCompactThresholdPercent: compactionConfig.autoCompactThresholdPercent });
        }
      }

      // 固定文件夹
      if (Array.isArray(allSettings.pinnedFolders)) {
        set({ pinnedFolders: allSettings.pinnedFolders });
      }

      // 知识库
      if (typeof allSettings.knowledgeEnabled === "boolean") {
        set({ knowledgeEnabled: allSettings.knowledgeEnabled });
      }

      set({
        isLoading: false,
        isConfigured: get().checkConfigured(),
      });
    } catch {
      set({ isLoading: false });
    }
  },

  saveSettings: async () => {
    const { providers, activeProviderId, permissionMode, showReasoning, language, theme, closeToTray, officeAutoCompactEnabled, dynamicArrayFunctionsEnabled, windowOpacity, compactionEnabled, autoCompactThresholdPercent, pinnedFolders } = get();

    await ipcApi.settings.set("aiProviders", providers);
    await ipcApi.settings.set("activeProvider", activeProviderId);
    await ipcApi.settings.set("permissionMode", permissionMode);
    await ipcApi.settings.set("showReasoning", showReasoning);
    await ipcApi.settings.set("language", language);
    await ipcApi.settings.set("theme", theme);
    await ipcApi.settings.set("closeToTray", closeToTray);
    await ipcApi.settings.set("officeAutoCompactEnabled", officeAutoCompactEnabled);
    await ipcApi.settings.set("dynamicArrayFunctionsEnabled", dynamicArrayFunctionsEnabled);
    await ipcApi.settings.set("windowOpacity", normalizeWindowOpacity(windowOpacity));
    // 上下文压缩配置（百分比制，main.ts 会根据当前模型换算为实际 token 阈值）
    const existingCompaction = await ipcApi.settings.get("compactionConfig") as Record<string, unknown> | null;
    await ipcApi.settings.set("compactionConfig", {
      ...(existingCompaction && typeof existingCompaction === "object" ? existingCompaction : {}),
      enabled: compactionEnabled,
      autoCompactThresholdPercent,
    });
    // 固定文件夹
    await ipcApi.settings.set("pinnedFolders", pinnedFolders);
  },

  setActiveProvider: (id: string) => {
    set({ activeProviderId: id, isConfigured: get().checkConfigured() });
    savePartial(["activeProviderId"], get);
  },

  updateProvider: (id: string, patch: Partial<AiProviderConfig>) => {
    set((s) => {
      const nextProvider = normalizeProviderConfig({ ...s.providers[id], ...patch });
      return {
        providers: {
          ...s.providers,
          [id]: nextProvider,
        },
        isConfigured: get().checkConfigured(),
      };
    });
    savePartial(["providers"], get);
  },

  addProvider: (config: AiProviderConfig) => {
    set((s) => {
      const normalizedConfig = normalizeProviderConfig(config);
      const newProviders = { ...s.providers, [normalizedConfig.id]: normalizedConfig };
      const newActiveId = s.activeProviderId || normalizedConfig.id;
      return {
        providers: newProviders,
        activeProviderId: newActiveId,
        isConfigured: !!(normalizedConfig.apiKey && normalizedConfig.baseUrl && (normalizedConfig.model || normalizedConfig.defaultModel)),
      };
    });
    savePartial(["providers", "activeProviderId"], get);
  },

  removeProvider: (id: string) => {
    set((s) => {
      const { [id]: _, ...rest } = s.providers;
      const remainingIds = Object.keys(rest);
      const newActiveId = s.activeProviderId === id
        ? (remainingIds.length > 0 ? remainingIds[0] : "")
        : s.activeProviderId;
      return {
        providers: rest,
        activeProviderId: newActiveId,
      };
    });
    set({ isConfigured: get().checkConfigured() });
    savePartial(["providers", "activeProviderId"], get);
  },

  setProviderModels: (id: string, models: string[]) => {
    set((s) => ({
      providers: {
        ...s.providers,
        [id]: { ...s.providers[id], models },
      },
    }));
    savePartial(["providers"], get);
  },

  generateId: () => {
    return `provider_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
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
    if (!activeProviderId) return false;
    const provider = providers[activeProviderId];
    return !!(provider?.apiKey && provider?.baseUrl && (provider?.model || provider?.defaultModel));
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
