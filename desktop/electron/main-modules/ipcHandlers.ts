/**
 * IPC 处理器 — 所有 IPC handle 注册
 *
 * 从 main.ts 提取，按功能域拆分：
 * - 应用信息 & 数据路径
 * - 窗口行为
 * - 设置读写
 * - Excel 连接
 * - 文件对话框
 * - 文件夹文件列表
 * - 文件读取
 * - 沙箱配置
 * - AI 模型列表 & 测试
 *
 * Agent 会话、线程、工具定义和知识库 IPC 已迁入 agent/interaction。
 */

import { dialog, shell, BrowserWindow } from "electron";
import type { AgentLoop } from "../agent/core/agentLoop";
import {
  ensureKnowledgeRuntime,
  refreshKnowledgeRuntime,
  type AgentLoopManager,
} from "../agent/runtime/agentRuntime";
import { inspectExcelWorkbookForIpc, readExcelRangeForIpc } from "./excelIpcOperations";
import { getOrCreateExcelBridge } from "../agent/runtime/bridgeRegistry";
import type { ExcelConnectionBridge } from "../agent/tools/contracts/excel";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import {
  buildCompactionConfig,
  type SavedCompactionConfig,
} from "../agent/runtime/compactionRuntime";
import { setDynamicArrayFunctionsEnabled } from "../agent/runtime/agentGlobalSettings";
import { registerAgentIpcHandlers } from "../agent/interaction/ipcAgentHandlers";
import {
  getSettingsStore,
  getSettingForRenderer,
  getSettingsForRenderer,
  setSettingFromRenderer,
  getActiveDataPath,
  getActiveAIConfig,
  getRuntimeSettingValue,
  getSessionStoreInstance,
  getStateRuntimeStoreInstance,
  isDataMigrationInProgress,
  getAgentGraphStoreInstance,
  eraseUserData,
  exportUserData,
  migrateDataPath,
  applyWindowOpacity,
  applyWindowTheme,
} from "./settingsManager";
import {
  validateInput,
  AppLogInput,
  AppOpenExternalInput,
  AppOpenPathInput,
  LaunchOfficeApplicationInput,
  UpdateCheckInput,
  UpdateKindInput,
  ExcelReadRangeInput,
  ExcelSelectHostInput,
  ExcelWriteRangeInput,
  EraseUserDataInput,
  ExportUserDataInput,
  MigrateDataPathInput,
  SettingsGetInput,
  SettingsSetInput,
  SetAlwaysOnTopInput,
  WindowDisplayModeInput,
} from "../shared/ipcSchemas";
import { launchOfficeApplication } from "./officeProcessLauncher";
import { createLogger } from "../shared/logger";
import { assertAuthorizedPath, createPathAuthorizer } from "./ipcPathSecurity";
import {
  configureTrustedIpcSender,
  trustedIpcMain as ipcMain,
} from "../shared/trustedIpc";
import { registerOcrIpcHandler } from "./ipcOcrHandlers";
import { registerAiIpcHandlers } from "./ipcAiHandlers";
import { registerFileIpcHandlers } from "./ipcFileHandlers";
import { registerOfficeAutomationIpcHandlers } from "./ipcOfficeAutomationHandlers";
import { guardDataOperation } from "./dataMaintenance";
import {
  getWindowDisplayMode,
  setWindowDisplayMode,
  type WindowDisplayMode,
} from "./windowManager";
import {
  acknowledgeHotPatchHealth,
  applyDownloadedUpdate,
  checkForUpdates,
  downloadUpdate,
  getUpdateState,
} from "./updateManager";

const logger = createLogger("IPC");
const rendererLogger = createLogger("renderer");

// ============================================================
// Globals (assigned by registerIpcHandlers)
// ============================================================

let mainWindowRef: () => BrowserWindow | null = () => null;
let agentLoopsRef: () => AgentLoop[] = () => [];
let agentLoopManagerRef: () => AgentLoopManager | null = () => null;
let excelBridgeRef: () => ExcelConnectionBridge | null;
let wordBridgeRef: () => any = () => null;
let presentationBridgeRef: () => any = () => null;

export function setMainWindowRef(fn: () => BrowserWindow | null): void {
  mainWindowRef = fn;
}

export function setAgentLoopsRef(fn: () => AgentLoop[]): void {
  agentLoopsRef = fn;
}

export function setAgentLoopManagerRef(fn: () => AgentLoopManager | null): void {
  agentLoopManagerRef = fn;
}

export function setExcelBridgeRef(excel: () => ExcelConnectionBridge | null): void {
  excelBridgeRef = excel;
}

export function setOfficeBridgesRefs(word: () => any, presentation: () => any): void {
  wordBridgeRef = word;
  presentationBridgeRef = presentation;
}

// ============================================================
// IPC 处理器注册
// ============================================================

export function registerIpcHandlers(): void {
  configureTrustedIpcSender(mainWindowRef);
  setDynamicArrayFunctionsEnabled(getSettingsStore().get("dynamicArrayFunctionsEnabled"));
  const pathAuthorizer = createPathAuthorizer({
    getDataPath: getActiveDataPath,
    getPinnedFolders: () => {
      const folders = getSettingsStore().get("pinnedFolders") as
        Array<{ path?: unknown }> | undefined;
      return Array.isArray(folders)
        ? folders
            .map((folder) => (typeof folder.path === "string" ? folder.path : ""))
            .filter(Boolean)
        : [];
    },
  });

  registerAgentIpcHandlers({
    mainWindowRef,
    agentLoopManagerRef,
    getSessionStoreInstance,
    getStateRuntimeStoreInstance,
    getAgentGraphStoreInstance,
    ensureKnowledgeRuntime: () => ensureKnowledgeRuntime(
      getActiveAIConfig(),
      getActiveDataPath(),
      () => getRuntimeSettingValue("remoteDataProcessingEnabled") === true,
    ),
    isDataMigrationInProgress,
    pathAuthorizer,
  });

  // ---- 应用信息 ----
  registerOcrIpcHandler(pathAuthorizer, isDataMigrationInProgress);
  registerAiIpcHandlers(isDataMigrationInProgress);
  registerFileIpcHandlers({
    mainWindowRef,
    pathAuthorizer,
    getDataPath: getActiveDataPath,
    isDataMaintenanceInProgress: isDataMigrationInProgress,
  });
  registerOfficeAutomationIpcHandlers({
    getDataPath: getActiveDataPath,
    isDataMaintenanceInProgress: isDataMigrationInProgress,
  });

  ipcMain.handle("app:getDataPath", () => getActiveDataPath());

  ipcMain.handle("app:openPath", async (_event, targetPath: unknown) => {
    const validated = validateInput(AppOpenPathInput, targetPath);
    return await shell.openPath(assertAuthorizedPath(pathAuthorizer, validated));
  });

  ipcMain.handle("app:openExternal", async (_event, targetUrl: unknown) => {
    try {
      const validated = validateInput(AppOpenExternalInput, targetUrl);
      const url = new URL(validated);
      if (url.protocol !== "https:" && url.protocol !== "http:") {
        return `Unsupported URL protocol: ${url.protocol}`;
      }
      await shell.openExternal(url.toString());
      return "";
    } catch (error: any) {
      return error?.message || "Failed to open external URL";
    }
  });

  ipcMain.handle("app:launchOffice", async (_event, application: unknown) => {
    try {
      const validated = validateInput(LaunchOfficeApplicationInput, application);
      return await launchOfficeApplication(validated);
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : "启动程序失败",
      };
    }
  });

  ipcMain.handle("update:getState", () => getUpdateState());
  ipcMain.handle("update:ackHotPatchHealth", () => acknowledgeHotPatchHealth());
  ipcMain.handle("update:check", (_event, manual: unknown) => {
    return checkForUpdates(validateInput(UpdateCheckInput, manual) ?? true);
  });
  ipcMain.handle("update:download", (_event, kind: unknown) => {
    return downloadUpdate(validateInput(UpdateKindInput, kind));
  });
  ipcMain.handle("update:apply", () => applyDownloadedUpdate());

  // 转发渲染进程日志到主进程持久化
  ipcMain.handle("app:log", (_event, level: unknown, tag: unknown, message: unknown) => {
    if (isDataMigrationInProgress()) return;
    const { level: levelStr, tag: tagStr, message: msgStr } = validateInput(AppLogInput, {
      level,
      tag,
      message,
    });
    const logMsg = `[${tagStr}] ${msgStr}`;
    if (levelStr === "error") rendererLogger.error(logMsg);
    else if (levelStr === "warn") rendererLogger.warn(logMsg);
    else rendererLogger.info(logMsg);
  });

  ipcMain.handle("app:selectDataPath", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] };
    return await dialog.showOpenDialog(mw, {
      title: "选择数据存储目录",
      properties: ["openDirectory", "createDirectory"],
    });
  });

  ipcMain.handle("app:migrateDataPath", async (_event, targetPath: unknown) => {
    const validated = validateInput(MigrateDataPathInput, targetPath);
    return await migrateDataPath(validated);
  });

  ipcMain.handle("app:selectExportPath", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] };
    return await dialog.showOpenDialog(mw, {
      title: "选择空目录导出本地数据",
      properties: ["openDirectory", "createDirectory"],
    });
  });

  ipcMain.handle("app:exportUserData", async (_event, targetPath: unknown) => {
    const validated = validateInput(ExportUserDataInput, targetPath);
    return await exportUserData(validated);
  });

  ipcMain.handle("app:eraseUserData", async (_event, input: unknown) => {
    const validated = validateInput(EraseUserDataInput, input);
    return await eraseUserData(validated.confirmation);
  });

  // ---- 窗口行为 ----
  ipcMain.handle("window:getAlwaysOnTop", () => {
    return mainWindowRef()?.isAlwaysOnTop() ?? false;
  });

  ipcMain.handle("window:setAlwaysOnTop", (_event, enabled: unknown) => {
    const validated = validateInput(SetAlwaysOnTopInput, enabled);
    mainWindowRef()?.setAlwaysOnTop(validated);
    return mainWindowRef()?.isAlwaysOnTop() ?? false;
  });

  ipcMain.handle("window:getDisplayMode", () => {
    return getWindowDisplayMode();
  });

  ipcMain.handle("window:setDisplayMode", (_event, mode: unknown) => {
    const validated = validateInput(WindowDisplayModeInput, mode);
    return setWindowDisplayMode(mainWindowRef(), validated as WindowDisplayMode);
  });

  // ---- 设置相关 ----
  ipcMain.handle("settings:get", (_event, key: unknown) => {
    const validated = validateInput(SettingsGetInput, key);
    return getSettingForRenderer(validated);
  });

  ipcMain.handle("settings:set", guardDataOperation(isDataMigrationInProgress, async (_event, keyInput: unknown, valueInput: unknown) => {
    const [key, value] = validateInput(SettingsSetInput, [keyInput, valueInput]);
    const rendererValue = setSettingFromRenderer(key, value);

    if (key === "activeProvider" || key === "aiProviders") {
      for (const agent of agentLoopsRef()) {
        agent.updateAIConfig(getActiveAIConfig());
        const aiConfig = getActiveAIConfig();
        const contextWindowSize = aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW;
        agent.updateCompactionConfig(
          buildCompactionConfig({
            contextWindowSize,
            savedCompaction: getRuntimeSettingValue("compactionConfig") as
              SavedCompactionConfig | undefined,
          }),
        );
      }
      try {
        await refreshKnowledgeRuntime(
          getActiveAIConfig(),
          getActiveDataPath(),
          () => getRuntimeSettingValue("remoteDataProcessingEnabled") === true,
        );
      } catch (error) {
        logger.warn("刷新知识库运行时失败，设置已保存:", error);
      }
    }
    if (key === "permissionMode") {
      for (const agent of agentLoopsRef()) {
        agent.updatePermissionMode(value as "normal" | "auto_approve_safe" | "confirm_all");
      }
    }
    if (key === "closeToTray" && value === true) {
      // Tray creation handled by windowManager
    }
    if (key === "closeToTray" && value === false) {
      // Tray destruction handled by windowManager
    }
    if (key === "compactionConfig") {
      for (const agent of agentLoopsRef()) {
        const aiConfig = getActiveAIConfig();
        const contextWindowSize = aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW;
        agent.updateCompactionConfig(
          buildCompactionConfig({
            contextWindowSize,
            savedCompaction: getRuntimeSettingValue("compactionConfig") as
              SavedCompactionConfig | undefined,
          }),
        );
      }
    }
    if (key === "theme") {
      applyWindowTheme(mainWindowRef());
    }
    if (key === "windowOpacity") {
      applyWindowOpacity(mainWindowRef());
    }
    if (key === "dynamicArrayFunctionsEnabled") {
      setDynamicArrayFunctionsEnabled(value);
    }
    return rendererValue;
  }));

  ipcMain.handle("settings:getAll", () => {
    return getSettingsForRenderer();
  });

  // ---- Excel 连接状态 ----
  ipcMain.handle("excel:detectStatus", async () => {
    return await getExcelBridgeForIpc().detectStatus();
  });

  ipcMain.handle("excel:connect", async () => {
    return await getExcelBridgeForIpc().connect();
  });

  ipcMain.handle("excel:selectHost", async (_event, host: unknown) => {
    const validated = validateInput(ExcelSelectHostInput, host);
    const bridge = excelBridgeRef();
    if (!bridge) return { connected: false, host: "unknown", error: "Bridge not available" };
    return await bridge.selectHost(validated);
  });

  // ---- Word 连接状态 ----
  ipcMain.handle("word:detectStatus", async () => {
    const bridge = wordBridgeRef();
    if (!bridge) return { connected: false, host: "unknown" };
    return await bridge.detectStatus();
  });

  // ---- PowerPoint 连接状态 ----
  ipcMain.handle("ppt:detectStatus", async () => {
    const bridge = presentationBridgeRef();
    if (!bridge) return { connected: false, host: "unknown" };
    return await bridge.detectStatus();
  });

  ipcMain.handle("excel:getSelection", async () => {
    const bridge = excelBridgeRef();
    if (!bridge) throw new Error("Excel 桥接未初始化");
    return await bridge.getSelection();
  });

  ipcMain.handle("excel:getSelectionAddress", async () => {
    const bridge = excelBridgeRef();
    if (!bridge) throw new Error("Excel 桥接未初始化");
    return await bridge.getSelectionAddress();
  });

  ipcMain.handle(
    "excel:readRange",
    async (_event, sheetName: unknown, range: unknown, expand?: unknown) => {
      const validated = validateInput(ExcelReadRangeInput, { sheetName, range, expand });
      return readExcelRangeForIpc(
        excelBridgeRef(),
        validated.sheetName,
        validated.range,
        validated.expand,
      );
    },
  );

  ipcMain.handle("excel:inspectWorkbook", async () => {
    return inspectExcelWorkbookForIpc(excelBridgeRef());
  });

  ipcMain.handle(
    "excel:writeRange",
    async (_event, sheetName: unknown, range: unknown, values: unknown) => {
      const validated = validateInput(ExcelWriteRangeInput, { sheetName, range, values });
      const bridge = excelBridgeRef();
      if (!bridge) return { success: false, error: "Excel 未连接" };
      try {
        const data = await bridge.writeRange(validated.sheetName, validated.range, validated.values);
        return { success: true, data };
      } catch (err: any) {
        return { success: false, error: err.message };
      }
    },
  );
}

function getExcelBridgeForIpc(): ExcelConnectionBridge {
  return excelBridgeRef() ?? getOrCreateExcelBridge();
}
