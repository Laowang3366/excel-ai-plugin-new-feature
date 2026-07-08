/**
 * IPC API 抽象层
 *
 * 封装 window.electronAPI 的直接访问，提供：
 * 1. 集中管理所有 IPC 调用
 * 2. 测试时可注入 mock 实现（通过 createMockIpcApi）
 * 3. 运行时安全检查（electronAPI 可能不存在）
 *
 * 使用方式：
 *   import { ipcApi } from "../services/ipcApi";
 *   const result = await ipcApi.settings.getAll();
 */

import type { IIpcApi } from "./ipcApiTypes";
import { createKnowledgeIpcApi } from "./ipcKnowledgeApi";
import { createOfficeIpcApi } from "./ipcOfficeApi";
import { createThreadIpcApi } from "./ipcThreadApi";
export type { IIpcApi } from "./ipcApiTypes";
export { createMockIpcApi } from "./ipcApiMock";

// ============================================================
// 实现
// ============================================================

/** 获取底层 electronAPI，不存在时返回 null */
function getRaw(): IIpcApi | null {
  if (typeof window === "undefined") return null;
  return (window as any).electronAPI ?? null;
}

/**
 * 运行时 IPC API 实例。
 *
 * 所有方法在 electronAPI 不可用时返回安全的空值，
 * 调用方无需额外判空。
 */
export const ipcApi: IIpcApi = {
  app: {
    getDataPath: async () => {
      const raw = getRaw();
      if (!raw) throw new Error("IPC not available: app.getDataPath");
      return raw.app.getDataPath();
    },
    selectDataPath: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.app.selectDataPath();
    },
    migrateDataPath: async (targetPath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.app.migrateDataPath(targetPath);
    },
    openPath: async (targetPath) => {
      const raw = getRaw();
      if (!raw) return "";
      return raw.app.openPath(targetPath);
    },
    openExternal: async (targetUrl) => {
      const raw = getRaw();
      if (!raw) return "";
      return raw.app.openExternal(targetUrl);
    },
    /** 记录日志 → 主进程持久化（由 rendererLogger.ts 调用） */
    log: async (level, tag, message) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.app.log(level, tag, message);
    },
  },

  window: {
    getAlwaysOnTop: async () => {
      const raw = getRaw();
      if (!raw) return false;
      return raw.window.getAlwaysOnTop();
    },
    setAlwaysOnTop: async (enabled) => {
      const raw = getRaw();
      if (!raw) return false;
      return raw.window.setAlwaysOnTop(enabled);
    },
    getDisplayMode: async () => {
      const raw = getRaw();
      if (!raw?.window.getDisplayMode) return "normal";
      return raw.window.getDisplayMode();
    },
    setDisplayMode: async (mode) => {
      const raw = getRaw();
      if (!raw?.window.setDisplayMode) return "normal";
      return raw.window.setDisplayMode(mode);
    },
    onDisplayModeChanged: (callback) => {
      const raw = getRaw();
      if (!raw?.window.onDisplayModeChanged) return () => {};
      return raw.window.onDisplayModeChanged(callback);
    },
  },

  settings: {
    get: async (key) => {
      const raw = getRaw();
      if (!raw) return undefined;
      return raw.settings.get(key);
    },
    set: async (key, value) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.settings.set(key, value);
    },
    getAll: async () => {
      const raw = getRaw();
      if (!raw) return {};
      return raw.settings.getAll();
    },
  },

  ...createOfficeIpcApi(getRaw),

  agent: {
    startTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.startTurn(input);
    },
    continueTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.continueTurn(input);
    },
    enqueueTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.enqueueTurn(input);
    },
    interrupt: async (threadId?: string | null) => {
      const raw = getRaw();
      if (!raw) return { success: false };
      return raw.agent.interrupt(threadId);
    },
    onEvent: (callback) => {
      const raw = getRaw();
      if (!raw) return () => {};
      return raw.agent.onEvent(callback);
    },
    onStreamDelta: (callback) => {
      const raw = getRaw();
      if (!raw) return () => {};
      return raw.agent.onStreamDelta(callback);
    },
  },

  ...createThreadIpcApi(getRaw),

  dialog: {
    openFile: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openFile();
    },
    openImage: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openImage();
    },
    openFolder: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openFolder();
    },
  },

  file: {
    readAsBase64: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { error: "IPC not available" };
      return raw.file.readAsBase64(filePath);
    },
    trashFile: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.trashFile(filePath);
    },
    openFile: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.openFile(filePath);
    },
    copyPath: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false };
      return raw.file.copyPath(filePath);
    },
    revealInExplorer: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.revealInExplorer(filePath);
    },
    writeTempFile: async (data) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.writeTempFile(data);
    },
    getPathForFile: (file) => {
      const raw = getRaw();
      if (!raw?.file.getPathForFile) return "";
      return raw.file.getPathForFile(file);
    },
  },

  folder: {
    listFiles: async (folderPath) => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.folder.listFiles(folderPath);
    },
    listFilesBatch: async (folderPaths) => {
      const raw = getRaw();
      if (!raw) return {};
      if (raw.folder.listFilesBatch) return raw.folder.listFilesBatch(folderPaths);
      const entries = await Promise.all(
        folderPaths.map(async (folderPath) => {
          try {
            return [folderPath, await raw.folder.listFiles(folderPath)] as const;
          } catch {
            return [folderPath, []] as const;
          }
        })
      );
      return Object.fromEntries(entries);
    },
  },

  tools: {
    list: async () => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.tools.list();
    },
  },

  sandbox: {
    getConfig: async () => {
      const raw = getRaw();
      if (!raw) return { defaultRules: [], userRules: [], extraWritableRoots: [] };
      return raw.sandbox.getConfig();
    },
    setUserRules: async (rules) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.sandbox.setUserRules(rules);
    },
    setWritableRoots: async (roots) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.sandbox.setWritableRoots(roots);
    },
  },

  tool: {
    confirm: async (toolCallId, alwaysAllow) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.tool.confirm(toolCallId, alwaysAllow);
    },
    cancel: async (toolCallId) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.tool.cancel(toolCallId);
    },
  },

  ai: {
    listModels: async (baseUrl, apiKey, apiFormat) => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.ai.listModels(baseUrl, apiKey, apiFormat);
    },
    testConnection: async (baseUrl, apiKey, apiFormat, model) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.ai.testConnection(baseUrl, apiKey, apiFormat, model);
    },
  },

  stats: {
    getSummary: async () => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.stats.getSummary();
    },
  },

  ocr: {
    recognize: async (mode, filePaths) => {
      const raw = getRaw() as any;
      if (raw?.ocr?.recognize) {
        return raw.ocr.recognize(mode, filePaths);
      }
      return null;
    },
  },

  knowledge: createKnowledgeIpcApi(getRaw),

  /**
   * 激活管理（activation）
   *
   * 封装与主进程激活模块的 IPC 通信，包括：
   * - getStatus / checkValid — 查询当前激活状态及有效性校验
   * - activate — 使用卡密向许可证服务器发起激活
   * - clear — 清除本地激活凭证（反激活）
   * - getServerUrl / setServerUrl — 许可证服务器地址配置
   * - getMachineInfo — 获取本机设备信息（用于绑定展示）
   * - listDevices / unbindDevice — 已绑定的设备列表管理与解绑
   *
   * electronAPI 不可用时返回安全的兜底值（如 { activated: false }），
   * 上层 Store 据此进入未激活流程。
   */
  activation: {
    /** 查询当前激活状态（是否已激活、激活信息等） */
    getStatus: async () => {
      const raw = getRaw();
      if (!raw?.activation?.getStatus) return { activated: false };
      return raw.activation.getStatus();
    },
    /** 使用卡密向许可证服务器发起激活，返回 { success, error? } */
    activate: async (key, serverUrl) => {
      const raw = getRaw();
      if (!raw?.activation?.activate) return { success: false, error: "IPC not available" };
      return raw.activation.activate(key, serverUrl);
    },
    /** 清除本地激活凭证（反激活） */
    clear: async () => {
      const raw = getRaw();
      if (!raw?.activation?.clear) return { success: false };
      return raw.activation.clear();
    },
    /** 获取当前许可证服务器地址 */
    getServerUrl: async () => {
      const raw = getRaw();
      if (!raw?.activation?.getServerUrl) return "http://localhost:3456";
      return raw.activation.getServerUrl();
    },
    /** 设置许可证服务器地址 */
    setServerUrl: async (url) => {
      const raw = getRaw();
      if (!raw?.activation?.setServerUrl) return { success: false };
      return raw.activation.setServerUrl(url);
    },
    /** 校验当前激活是否仍有效（调用服务端验证接口） */
    checkValid: async () => {
      const raw = getRaw();
      if (!raw?.activation?.checkValid) return false;
      return raw.activation.checkValid();
    },
    /** 获取本机设备信息（机器标识和名称） */
    getMachineInfo: async () => {
      const raw = getRaw();
      if (!raw?.activation?.getMachineInfo) return { machineId: "", machineName: "" };
      return raw.activation.getMachineInfo();
    },
    /** 列出当前卡密已绑定的设备列表 */
    /** 列出当前卡密已绑定的设备列表 */
    listDevices: async () => {
      const raw = getRaw();
      if (!raw?.activation?.listDevices) return { success: false, error: "IPC not available" };
      return raw.activation.listDevices();
    },
    /** 解绑指定设备（释放该设备的激活名额） */
    unbindDevice: async (targetMachineId) => {
      const raw = getRaw();
      if (!raw?.activation?.unbindDevice) return { success: false, error: "IPC not available" };
      return raw.activation.unbindDevice(targetMachineId);
    },
  },
};
