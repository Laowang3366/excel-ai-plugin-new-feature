/**
 * Electron 主进程入口（精简版 — 仅协调各模块）
 *
 * 已拆分模块：
 * - main-modules/settingsManager.ts: 持久化配置管理
 * - main-modules/windowManager.ts: 窗口创建与托盘
 * - main-modules/ipcHandlers.ts: IPC 处理器
 * - agent/interaction/eventForwarder.ts: 事件转发与工具审批
 */

import { app, BrowserWindow } from "electron";
import * as path from "path";
import {
  getAgentLoop,
  getAgentLoopManager,
  getAgentLoops,
  getOrCreateAgentRuntime,
} from "./agent/runtime/agentRuntime";
import {
  disconnectOfficeBridges,
  getExcelBridge,
  getExcelScriptBridge,
  getExcelUiBridge,
  getVbaBridge,
  getWordBridge,
  getPresentationBridge,
} from "./agent/runtime/bridgeRegistry";
import {
  getSettingsStore,
  getActiveDataPath,
  getActiveAIConfig,
  getSessionStoreInstance,
  getStateRuntimeStoreInstance,
  closeStateRuntimeStore,
  setAgentLoopGetter,
  setAgentLoopsGetter,
} from "./main-modules/settingsManager";
import {
  createWindow,
  setIsQuitting,
} from "./main-modules/windowManager";
import {
  registerIpcHandlers,
  setMainWindowRef,
  setAgentLoopRef,
  setAgentLoopManagerRef,
  setAgentLoopsRef,
  setBridgesRefs,
  setOfficeBridgesRefs,
  applySandboxConfig,
} from "./main-modules/ipcHandlers";
import { requestToolApproval } from "./agent/interaction/eventForwarder";
import { configureLogDirectory, setupGlobalErrorHandlers } from "./shared/logger";

// ============================================================
// 全局错误捕获（尽早注册，确保所有异常写入日志）
// ============================================================
configureLogDirectory(path.join(getActiveDataPath(), "logs"));
setupGlobalErrorHandlers();

// ============================================================
// 全局状态
// ============================================================

let mainWindow: BrowserWindow | null = null;

// ==== 引用函数（供 IPC 处理器使用）====
setMainWindowRef(() => mainWindow);
setAgentLoopRef(() => getAgentLoop());
setAgentLoopsRef(() => getAgentLoops());
setAgentLoopManagerRef(() => getAgentLoopManager());
setBridgesRefs(
  () => getExcelBridge(),
  () => getVbaBridge(),
  () => getExcelScriptBridge(),
  () => getExcelUiBridge(),
);
setOfficeBridgesRefs(
  () => getWordBridge(),
  () => getPresentationBridge(),
);
// 给 settingsManager 提供 AgentLoop 引用（数据迁移后刷新 SessionStore）
setAgentLoopGetter(() => getAgentLoop());
setAgentLoopsGetter(() => getAgentLoops());

// 窗口重建函数（用于关闭到托盘后重新创建）
function recreateMainWindow(): BrowserWindow {
  const mw = createWindow(recreateMainWindow, (win) => { mainWindow = win; });
  mainWindow = mw;
  return mw;
}

// ============================================================
// 应用生命周期
// ============================================================

app.whenReady().then(async () => {
  getSessionStoreInstance(); // 提前初始化 SessionStore
  await getOrCreateAgentRuntime({
    getActiveAIConfig,
    getActiveDataPath,
    getSettingsValue: (key) => getSettingsStore().get(key as any),
    getSessionStoreInstance,
    getStateRuntimeStoreInstance,
    requestToolApproval: (params) => requestToolApproval(() => mainWindow, params),
  }); // 提前初始化 Agent（含 Office bridge + RAG）
  registerIpcHandlers();
  applySandboxConfig();    // 把 electron-store 中的用户规则热更新到沙箱单例
  recreateMainWindow();     // 创建窗口并保存引用

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      recreateMainWindow();
    } else if (mainWindow) {
      mainWindow.show();
      mainWindow.setSkipTaskbar(false);
      mainWindow.focus();
    }
  });
}).catch((err) => {
  console.error("[main] Fatal startup error:", err);
  // 如果窗口未创建，尝试创建并显示错误页面
  if (!mainWindow) {
    const mw = new BrowserWindow({
      width: 600,
      height: 400,
      title: "启动失败",
      webPreferences: { contextIsolation: true, nodeIntegration: false },
    });
    mw.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(
      `<!DOCTYPE html><html><body style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;background:#1e293b;color:#f8fafc">
        <div style="text-align:center"><h2>🚨 启动失败</h2><p style="color:#94a3b8">${err instanceof Error ? err.message : String(err)}</p>
        <pre style="font-size:12px;color:#64748b;max-width:500px;overflow:auto">${err instanceof Error ? (err.stack || '') : ''}</pre></div></body></html>`
    )}`);
    mw.show();
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", async () => {
  setIsQuitting(true);
  await getSessionStoreInstance().flushRolloutWrites();
  await closeStateRuntimeStore();
  await disconnectOfficeBridges();
});
