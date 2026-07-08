/**
 * 激活状态 Store
 *
 * 管理激活状态、服务器地址，提供激活/反激活操作。
 */

import { create } from "zustand";
import { ipcApi } from "../services/ipcApi";
import { logError } from "../utils/rendererLogger";

const TAG = "ActivationStore";

export interface ActivationStoreState {
  /** 是否已激活 */
  activated: boolean;
  /** 是否正在加载 */
  isLoading: boolean;
  /** 激活信息 */
  info: {
    key: string;
    machineId: string;
    activatedAt: string;
    lastVerifiedAt: string;
    expiresAt: string | null;
    serverUrl: string;
  } | null;
  /** 服务器地址 */
  serverUrl: string;
  /** 是否显示激活弹窗 */
  showActivationDialog: boolean;

  /** 加载激活状态 */
  loadStatus: () => Promise<void>;
  /** 执行激活 */
  activate: (key: string) => Promise<{ success: boolean; error?: string }>;
  /** 清除激活 */
  deactivate: () => Promise<void>;
  /** 设置服务器地址 */
  setServerUrl: (url: string) => Promise<void>;
  /** 设置弹窗显示 */
  setShowActivationDialog: (show: boolean) => void;
}

export const useActivationStore = create<ActivationStoreState>((set, get) => ({
  activated: false,
  isLoading: true,
  info: null,
  serverUrl: "http://localhost:3456",
  showActivationDialog: false,

  loadStatus: async () => {
    try {
      const [status, serverUrl, machineInfo] = await Promise.all([
        ipcApi.activation.getStatus(),
        ipcApi.activation.getServerUrl(),
        ipcApi.activation.getMachineInfo(),
      ]);

      const valid = status.activated ? await ipcApi.activation.checkValid() : false;

      set({
        activated: status.activated && valid,
        info: status.activated && status.info ? {
          key: status.info.key,
          machineId: status.info.machineId,
          activatedAt: status.info.activatedAt,
          lastVerifiedAt: status.info.lastVerifiedAt,
          expiresAt: status.info.expiresAt,
          serverUrl: status.info.serverUrl,
        } : null,
        serverUrl: serverUrl || "http://localhost:3456",
        showActivationDialog: !(status.activated && valid),
        isLoading: false,
      });
    } catch (err) {
      logError(TAG, "加载激活状态失败", err);
      set({ isLoading: false, showActivationDialog: true });
    }
  },

  activate: async (key: string) => {
    const { serverUrl } = get();
    try {
      const result = await ipcApi.activation.activate(key, serverUrl);
      if (result.success) {
        await get().loadStatus();
      }
      return result;
    } catch (err: any) {
      return { success: false, error: err.message || "激活失败" };
    }
  },

  deactivate: async () => {
    try {
      await ipcApi.activation.clear();
      set({ activated: false, info: null, showActivationDialog: true });
    } catch (err) {
      logError(TAG, "清除激活状态失败", err);
    }
  },

  setServerUrl: async (url: string) => {
    try {
      await ipcApi.activation.setServerUrl(url);
      set({ serverUrl: url });
    } catch (err) {
      logError(TAG, "设置服务器地址失败", err);
    }
  },

  setShowActivationDialog: (show: boolean) => {
    set({ showActivationDialog: show });
  },
}));
