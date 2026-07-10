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
import { configureLogDirectory, createLogger, setupGlobalErrorHandlers } from "./shared/logger";

const mainLogger = createLogger("main");

// ============================================================
// 全局错误捕获（尽早注册，确保所有异常写入日志）
// ============================================================
configureLogDirectory(path.join(getActiveDataPath(), "logs"));
setupGlobalErrorHandlers();

// ============================================================
// 全局状态
// ============================================================

let mainWindow: BrowserWindow | null = null;

// ==== Getter closures for module references ====
// IPC handlers are registered before Agent Runtime and Window are fully
// initialized (see initialization order below). Passing getter closures
// (() => T) instead of direct values ensures handlers always read the
// latest reference at invocation time, avoiding initialization-order
// dependency issues.
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

// ---- 启动初始化序列（顺序敏感）----
// 1. SessionStore 提前初始化，确保后续模块可读写持久化数据
// 2. Agent Runtime 初始化（含 Office bridge + RAG），这是核心业务引擎
// 3. 注册 IPC 处理器，使渲染进程可调用主进程能力
// 4. 将用户配置的沙箱规则热更新到沙箱单例
// 5. 创建窗口并保存引用，此时 UI 就绪
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
// ---- 启动失败处理 ----
// 如果启动过程中任何一步抛出异常，捕获并记录日志。
// 若窗口尚未创建，则创建一个简易错误窗口展示异常信息，
// 确保用户能看到错误提示而不是白屏。
}).catch((err) => {
  mainLogger.error("Fatal startup error", err instanceof Error ? { message: err.message, stack: err.stack } : { error: String(err) });
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

// ---- 应用关闭前清理（顺序敏感）----
// 1. 标记正在退出，阻止窗口管理器的误操作（如最小化到托盘）
// 2. 刷写 SessionStore 中积压的滚动写入
// 3. 关闭状态运行时存储
// 4. 断开 Office bridge 连接（Excel/WPS 等 COM 对象）
app.on("before-quit", async () => {
  setIsQuitting(true);
  await getSessionStoreInstance().flushRolloutWrites();
  await closeStateRuntimeStore();
  await disconnectOfficeBridges();
});
