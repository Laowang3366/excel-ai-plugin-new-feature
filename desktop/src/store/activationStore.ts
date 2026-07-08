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
  /** 是否正在加载中（首次启动时从主进程获取状态） */
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
  /**
   * 是否显示激活弹窗
   *
   * 应用启动时若未激活或激活已失效，该值自动为 true。
   * 激活成功后通过 setShowActivationDialog(false) 关闭弹窗。
   */
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

/**
 * 使用 Zustand 创建的激活状态 Store
 *
 * Zustand 是一种轻量级状态管理库，通过 create<T>((set, get) => ...) 创建全局 store。
 * - set() 用于更新状态片段，触发组件重渲染
 * - get() 用于在 action 中读取当前状态快照，避免闭包过期问题
 *
 * 状态机关系说明：
 * - isLoading = true  → 应用启动中，正在从主进程查询激活状态
 * - activated = false + isLoading = false → 未激活，showActivationDialog 为 true
 * - activated = true + info 有值 → 已激活，showActivationDialog 为 false
 * - 激活校验失败（checkValid 返回 false）→ 自动令 activated = false 并弹窗
 */
export const useActivationStore = create<ActivationStoreState>((set, get) => ({
  activated: false,
  isLoading: true,
  info: null,
  serverUrl: "http://localhost:3456",
  showActivationDialog: false,

  /**
   * 从主进程加载激活状态
   *
   * 并发调用三个 IPC 接口获取全量状态，然后校验激活是否仍然有效。
   * 若校验失败或接口异常，将弹窗提示用户重新激活。
   *
   * 错误处理策略：
   * - 任意 IPC 调用抛出异常 → 记录错误日志，展示激活弹窗
   * - checkValid 返回 false → 标记为未激活，展示弹窗
   */
  loadStatus: async () => {
    try {
      const [status, serverUrl, machineInfo] = await Promise.all([
        ipcApi.activation.getStatus(),
        ipcApi.activation.getServerUrl(),
        ipcApi.activation.getMachineInfo(),
      ]);

      const valid = status.activated ? await ipcApi.activation.checkValid() : false;

      /**
       * 根据校验结果更新状态：
       * - 激活有效 → 填充 info，隐藏弹窗
       * - 激活无效 → 清空 info，展示弹窗（强反馈，要求用户重新激活）
       */
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
      /**
       * IPC 调用出错的降级处理：
       * 标记加载完成并弹出激活框，让用户至少能看到界面而不是卡死在 loading 状态。
       */
      logError(TAG, "加载激活状态失败", err);
      set({ isLoading: false, showActivationDialog: true });
    }
  },

  /**
   * 执行激活操作
   *
   * 调用主进程 IPC 完成与服务端的卡密校验。
   * 激活成功后自动重新加载全量状态（loadStatus 会再次核验）。
   *
   * 错误处理：捕获所有异常并返回结构化的失败结果，UI 层据此展示错误提示。
   */
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

  /**
   * 反激活（清除本地激活状态）
   *
   * 调用 IPC 清除本地存储的激活凭据，同时置空 store 并弹出激活框，
   * 使应用回到未激活的初始状态。
   * 注意：反激活通常由用户主动触发，不校验服务端是否可达。
   */
  deactivate: async () => {
    try {
      await ipcApi.activation.clear();
      set({ activated: false, info: null, showActivationDialog: true });
    } catch (err) {
      logError(TAG, "清除激活状态失败", err);
    }
  },

  /**
   * 设置许可证服务器地址
   *
   * 将地址持久化到主进程配置中，同时更新 store 中的 serverUrl。
   * 若主进程写入失败（例如配置目录无权限），仅记录错误而不阻塞 UI。
   */
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
