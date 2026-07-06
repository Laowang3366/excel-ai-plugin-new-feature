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

import { ipcMain, dialog, shell, clipboard, BrowserWindow } from "electron";
import * as path from "path";
import * as fs from "fs";
import { AgentLoop } from "../agent/core/agentLoop";
import { ensureKnowledgeRuntime, refreshKnowledgeRuntime, type AgentLoopManager } from "../agent/runtime/agentRuntime";
import { getOrCreateExcelBridge } from "../agent/runtime/bridgeRegistry";
import type { ExcelConnectionBridge } from "../agent/tools/contracts/excel";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import { buildCompactionConfig, type SavedCompactionConfig } from "../agent/runtime/compactionRuntime";
import { setDynamicArrayFunctionsEnabled } from "../agent/runtime/agentGlobalSettings";
import { registerAgentIpcHandlers } from "../agent/interaction/ipcAgentHandlers";
import {
  getSettingsStore,
  getActiveDataPath,
  getActiveAIConfig,
  getSessionStoreInstance,
  getStateRuntimeStoreInstance,
  getAgentGraphStoreInstance,
  migrateDataPath,
  applyWindowOpacity,
  applyWindowTheme,
} from "./settingsManager";
import {
  DEFAULT_RULES,
  setUserRules,
  setExtraWritableRoots,
  type PrefixRule,
} from "../agent/security/sandbox";
import {
  validateInput,
  AppOpenExternalInput,
  AppOpenPathInput,
  AiListModelsInput,
  AiTestConnectionInput,
  ExcelReadRangeInput,
  ExcelSelectHostInput,
  ExcelWriteRangeInput,
  FilePathInput,
  FileWriteTempFileInput,
  FolderPathInput,
  MigrateDataPathInput,
  SandboxUserRulesInput,
  SandboxWritableRootsInput,
  SettingsGetInput,
  SettingsSetInput,
  SetAlwaysOnTopInput,
  WindowDisplayModeInput,
} from "../shared/ipcSchemas";
import { createLogger } from "../shared/logger";
import { assertAuthorizedPath, createPathAuthorizer } from "./ipcPathSecurity";
import { registerOcrIpcHandler } from "./ipcOcrHandlers";
import {
  getWindowDisplayMode,
  setWindowDisplayMode,
  type WindowDisplayMode,
} from "./windowManager";

const logger = createLogger("IPC");

// ============================================================
// Globals (assigned by registerIpcHandlers)
// ============================================================

let mainWindowRef: () => BrowserWindow | null = () => null;
let agentLoopRef: () => AgentLoop | null;
let agentLoopsRef: () => AgentLoop[] = () => [];
let agentLoopManagerRef: () => AgentLoopManager | null = () => null;
let excelBridgeRef: () => ExcelConnectionBridge | null;
let vbaBridgeRef: () => any;
let scriptBridgeRef: () => any;
let uiBridgeRef: () => any;
let wordBridgeRef: () => any = () => null;
let presentationBridgeRef: () => any = () => null;

export function setMainWindowRef(fn: () => BrowserWindow | null): void {
  mainWindowRef = fn;
}

export function setAgentLoopRef(fn: () => AgentLoop | null): void {
  agentLoopRef = fn;
}

export function setAgentLoopsRef(fn: () => AgentLoop[]): void {
  agentLoopsRef = fn;
}

export function setAgentLoopManagerRef(fn: () => AgentLoopManager | null): void {
  agentLoopManagerRef = fn;
}

export function setBridgesRefs(
  excel: () => ExcelConnectionBridge | null,
  vba: () => any,
  script: () => any,
  ui: () => any
): void {
  excelBridgeRef = excel;
  vbaBridgeRef = vba;
  scriptBridgeRef = script;
  uiBridgeRef = ui;
}

export function setOfficeBridgesRefs(
  word: () => any,
  presentation: () => any
): void {
  wordBridgeRef = word;
  presentationBridgeRef = presentation;
}

// ============================================================
// IPC 处理器注册
// ============================================================

export function registerIpcHandlers(): void {
  setDynamicArrayFunctionsEnabled(getSettingsStore().get("dynamicArrayFunctionsEnabled"));
  const pathAuthorizer = createPathAuthorizer({
    getDataPath: getActiveDataPath,
    getPinnedFolders: () => {
      const folders = getSettingsStore().get("pinnedFolders") as Array<{ path?: unknown }> | undefined;
      return Array.isArray(folders)
        ? folders.map((folder) => typeof folder.path === "string" ? folder.path : "").filter(Boolean)
        : [];
    },
    getExtraRoots: () => {
      const roots = getSettingsStore().get("sandboxExtraWritableRoots") as unknown;
      return Array.isArray(roots)
        ? roots.filter((root): root is string => typeof root === "string" && root.trim().length > 0)
        : [];
    },
  });

  registerAgentIpcHandlers({
    mainWindowRef,
    agentLoopRef,
    agentLoopManagerRef,
    getSessionStoreInstance,
    getStateRuntimeStoreInstance,
    getAgentGraphStoreInstance,
    ensureKnowledgeRuntime: () => ensureKnowledgeRuntime(getActiveAIConfig(), getActiveDataPath()),
  });

  // ---- 应用信息 ----
  registerOcrIpcHandler(pathAuthorizer);

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
    return getSettingsStore().get(validated);
  });

  ipcMain.handle("settings:set", async (_event, keyInput: unknown, valueInput: unknown) => {
    const [key, value] = validateInput(SettingsSetInput, [keyInput, valueInput]);
    const store = getSettingsStore();
    store.set(key, value);

    if (key === "activeProvider" || key === "aiProviders") {
      for (const agent of agentLoopsRef()) {
        agent.updateAIConfig(getActiveAIConfig());
        const aiConfig = getActiveAIConfig();
        const contextWindowSize = aiConfig.contextWindowSize || DEFAULT_CONTEXT_WINDOW;
        agent.updateCompactionConfig(buildCompactionConfig({
          contextWindowSize,
          savedCompaction: store.get("compactionConfig") as SavedCompactionConfig | undefined,
        }));
      }
      try {
        await refreshKnowledgeRuntime(getActiveAIConfig(), getActiveDataPath());
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
        agent.updateCompactionConfig(buildCompactionConfig({
          contextWindowSize,
          savedCompaction: value as SavedCompactionConfig | undefined,
        }));
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
  });

  ipcMain.handle("settings:getAll", () => {
    return getSettingsStore().store;
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

  ipcMain.handle("excel:readRange", async (_event, sheetName: unknown, range: unknown, expand?: unknown) => {
    const validated = validateInput(ExcelReadRangeInput, { sheetName, range, expand });
    const bridge = excelBridgeRef();
    if (!bridge) return { values: [[]] };
    try { return await bridge.readRange(validated.sheetName, validated.range, validated.expand); }
    catch { return { values: [[]] }; }
  });

  ipcMain.handle("excel:inspectWorkbook", async () => {
    const bridge = excelBridgeRef();
    if (!bridge) return null;
    try { return await bridge.inspectWorkbook(); }
    catch { return null; }
  });

  ipcMain.handle("excel:writeRange", async (_event, sheetName: unknown, range: unknown, values: unknown) => {
    const validated = validateInput(ExcelWriteRangeInput, { sheetName, range, values });
    const bridge = excelBridgeRef();
    if (!bridge) return { success: false, error: "Excel 未连接" };
    try {
      await bridge.writeRange(validated.sheetName, validated.range, validated.values);
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  // ---- 文件对话框 ----
  ipcMain.handle("dialog:openFile", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      properties: ["openFile"],
      filters: [
        { name: "Documents", extensions: ["xlsx", "xls", "csv", "doc", "docx", "ppt", "pptx", "json", "txt", "pdf", "md"] },
        { name: "All Files", extensions: ["*"] },
      ],
    });
    result.filePaths.forEach((filePath) => pathAuthorizer.authorizePath(filePath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("dialog:openImage", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      properties: ["openFile"],
      filters: [
        { name: "Images", extensions: ["png", "jpg", "jpeg", "webp", "bmp", "gif"] },
      ],
    });
    result.filePaths.forEach((filePath) => pathAuthorizer.authorizePath(filePath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("dialog:openFolder", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      title: "选择文件夹",
      properties: ["openDirectory"],
    });
    result.filePaths.forEach((folderPath) => pathAuthorizer.authorizeRoot(folderPath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  // ---- 文件夹文件列表 ----
  ipcMain.handle("folder:listFiles", async (_event, folderPath: unknown) => {
    const validated = validateInput(FolderPathInput, folderPath);
    try {
      const authorizedFolderPath = assertAuthorizedPath(pathAuthorizer, validated);
      const officeExts = new Set([".xlsx", ".xls", ".csv", ".doc", ".docx", ".ppt", ".pptx"]);
      const entries = await fs.promises.readdir(authorizedFolderPath, { withFileTypes: true });
      const files = entries
        .filter((e) => e.isFile() && officeExts.has(path.extname(e.name).toLowerCase()))
        .map((e) => {
          const fullPath = path.join(authorizedFolderPath, e.name);
          pathAuthorizer.authorizePath(fullPath);
          return { filePath: fullPath, fileName: e.name };
        });
      const results = await Promise.all(
        files.map(async (f) => {
          try {
            const stat = await fs.promises.stat(f.filePath);
            return { ...f, size: stat.size, lastModified: stat.mtimeMs };
          } catch {
            return { ...f, size: 0, lastModified: 0 };
          }
        })
      );
      results.sort((a, b) => a.fileName.localeCompare(b.fileName));
      return results;
    } catch {
      return [];
    }
  });

  // ---- 文件读取 ----
  ipcMain.on("file:authorizePathSync", (event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      pathAuthorizer.authorizePath(validated);
      event.returnValue = { success: true };
    } catch (err: any) {
      event.returnValue = { success: false, error: err?.message || "授权路径失败" };
    }
  });

  ipcMain.handle("file:readAsBase64", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      const authorizedFilePath = assertAuthorizedPath(pathAuthorizer, validated);
      const buffer = await fs.promises.readFile(authorizedFilePath);
      const ext = path.extname(authorizedFilePath).toLowerCase().replace(".", "");
      const mimeMap: Record<string, string> = {
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", webp: "image/webp", bmp: "image/bmp",
        pdf: "application/pdf", csv: "text/csv", json: "application/json",
        txt: "text/plain", md: "text/markdown", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        xls: "application/vnd.ms-excel",
        doc: "application/msword",
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ppt: "application/vnd.ms-powerpoint",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      };
      const mimeType = mimeMap[ext] || "application/octet-stream";
      return {
        data: buffer.toString("base64"),
        mimeType,
        fileName: path.basename(authorizedFilePath),
        size: buffer.length,
      };
    } catch (err: any) {
      return { error: err.message };
    }
  });

  // ---- 临时文件操作（截图粘贴等） ----
  ipcMain.handle("file:writeTempFile", async (_event, data: unknown) => {
    try {
      const input = validateInput(FileWriteTempFileInput, data);
      const prefix = input.prefix?.replace(/[^a-zA-Z0-9_-]/g, "") || "clipboard";
      const suffix = input.suffix?.replace(/[^a-zA-Z0-9.]/g, "") || ".png";
      const tmpDir = await import("os").then((os) => os.tmpdir());
      const fileName = `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}${suffix}`;
      const filePath = path.join(tmpDir, fileName);
      const buffer = Buffer.from(input.data, "base64");
      await fs.promises.writeFile(filePath, buffer);
      pathAuthorizer.authorizePath(filePath);
      return { success: true, filePath };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  // ---- 文件操作（回收站/打开/复制路径/显示） ----
  ipcMain.handle("file:trashFile", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      await shell.trashItem(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:openFile", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      const result = await shell.openPath(assertAuthorizedPath(pathAuthorizer, validated));
      if (result) return { success: false, error: result };
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:copyPath", (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      clipboard.writeText(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch (err: any) {
      return { success: false };
    }
  });

  ipcMain.handle("file:revealInExplorer", (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      shell.showItemInFolder(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  // ---- 沙箱策略配置 ----
  // 暴露给设置页：展示生效规则、用户自定义规则、可写根
  ipcMain.handle("sandbox:getConfig", () => {
    const store = getSettingsStore();
    return {
      defaultRules: getSandboxDefaultRulesForUI(),
      userRules: store.get("sandboxUserRules") as PrefixRule[] | undefined ?? [],
      extraWritableRoots: store.get("sandboxExtraWritableRoots") as string[] | undefined ?? [],
    };
  });

  ipcMain.handle("sandbox:setUserRules", (_event, rules: unknown) => {
    const validated = validateInput(SandboxUserRulesInput, rules);
    const normalized = normalizeUserRules(validated);
    if (normalized.error) {
      return { success: false, error: normalized.error };
    }
    getSettingsStore().set("sandboxUserRules", normalized.rules);
    applySandboxConfig();
    return { success: true };
  });

  ipcMain.handle("sandbox:setWritableRoots", (_event, roots: unknown) => {
    const clean = validateInput(SandboxWritableRootsInput, roots)
      .map((root) => root.trim())
      .filter(Boolean);
    getSettingsStore().set("sandboxExtraWritableRoots", clean);
    applySandboxConfig();
    return { success: true };
  });

  // ---- AI 模型列表 + 测试 ----
  ipcMain.handle("ai:listModels", async (_event, baseUrl: unknown, apiKey: unknown, apiFormat: unknown) => {
    const validated = validateInput(AiListModelsInput, { baseUrl, apiKey, apiFormat });
    try {
      if (validated.apiFormat === "anthropic") return [];
      const url = validated.baseUrl.endsWith("/models")
        ? validated.baseUrl
        : `${validated.baseUrl.replace(/\/+$/, "")}/models`;
      const response = await fetch(url, {
        method: "GET",
        headers: {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        },
        signal: AbortSignal.timeout(15000),
      });
      if (!response.ok) {
        const errorText = await response.text().catch(() => "");
        logger.error(`[ai:listModels] HTTP ${response.status}: ${errorText}`);
        return [];
      }
      const data = (await response.json()) as any;
      if (Array.isArray(data?.data)) {
        return data.data
          .map((m: any) => m.id || m.name || "")
          .filter((id: string) => id.length > 0)
          .sort();
      }
      if (Array.isArray(data)) {
        return data
          .map((m: any) => (typeof m === "string" ? m : m.id || m.name || ""))
          .filter((id: string) => id.length > 0)
          .sort();
      }
      return [];
    } catch (err: any) {
      logger.error("[ai:listModels] Error:", err?.message || err);
      return [];
    }
  });

  ipcMain.handle("ai:testConnection", async (_event, baseUrl: unknown, apiKey: unknown, apiFormat: unknown, model: unknown) => {
    const validated = validateInput(AiTestConnectionInput, { baseUrl, apiKey, apiFormat, model });
    const startTime = Date.now();
    try {
      let url: string;
      let body: any;
      let headers: Record<string, string>;

      if (validated.apiFormat === "anthropic") {
        // 与 anthropicClient.ts:147 保持一致：固定追加 /messages
        url = `${validated.baseUrl.replace(/\/+$/, "")}/messages`;
        headers = {
          "x-api-key": validated.apiKey,
          "anthropic-version": "2023-06-01",
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "claude-sonnet-4-20250514", max_tokens: 1, messages: [{ role: "user", content: "Hi" }] };
      } else if (validated.apiFormat === "responses") {
        // OpenAI Responses 格式 — /responses endpoint
        url = `${validated.baseUrl.replace(/\/+$/, "")}/responses`;
        headers = {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "gpt-4o", input: "Hi", max_output_tokens: 1 };
      } else {
        // Chat Completions 格式（默认）— 与 openaiCompatibleClient.ts:123 保持一致
        url = `${validated.baseUrl.replace(/\/+$/, "")}/chat/completions`;
        headers = {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "gpt-4o", max_tokens: 1, messages: [{ role: "user", content: "Hi" }] };
      }

      const response = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(30000),
      });

      const latency = Date.now() - startTime;
      if (response.ok) {
        return { success: true, latency };
      } else {
        const errorText = await response.text().catch(() => "");
        let errorMessage = `HTTP ${response.status}`;
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error?.message || errorJson.message || errorMessage;
        } catch { /* use raw error text */ }
        return { success: false, error: errorMessage, latency };
      }
    } catch (err: any) {
      const latency = Date.now() - startTime;
      return { success: false, error: err?.message || "连接失败", latency };
    }
  });
}

// ============================================================
// 沙箱策略配置辅助
// ============================================================

/** 给 UI 展示的默认规则（只读副本） */
function getSandboxDefaultRulesForUI(): PrefixRule[] {
  // 返回深拷贝以防止渲染进程误改影响主进程单例
  return JSON.parse(JSON.stringify(DEFAULT_RULES));
}

interface NormalizedRulesResult {
  rules: PrefixRule[];
  error?: string;
}

/** 校验用户输入的规则，规范成 PrefixRule[] */
function normalizeUserRules(input: unknown[]): NormalizedRulesResult {
  const rules: PrefixRule[] = [];
  for (let i = 0; i < input.length; i++) {
    const raw = input[i];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return { rules, error: `规则 #${i} 必须为对象` };
    }
    const r = raw as Record<string, unknown>;
    const pat = r.pattern;
    const firstFromPat = Array.isArray(pat) && pat.length > 0 ? pat[0] : undefined;
    const first = r.first ?? firstFromPat;
    if (typeof first !== "string" || first.length === 0) {
      return { rules, error: `规则 #${i} 缺少 first 或 first 非字符串` };
    }
    const decision = r.decision;
    if (decision !== "allow" && decision !== "prompt" && decision !== "forbidden") {
      return { rules, error: `规则 #${i} decision 必须为 allow/prompt/forbidden` };
    }
    const rest: PrefixRule["rest"] = [];
    const tailPattern: unknown[] = Array.isArray(pat) ? pat.slice(1) : Array.isArray(r.rest) ? r.rest as unknown[] : [];
    for (const tok of tailPattern) {
      if (Array.isArray(tok)) {
        if (tok.length === 0) return { rules, error: `规则 #${i} alts 不能为空` };
        rest.push(tok.length === 1
          ? { kind: "single", value: tok[0] as string }
          : { kind: "alts", values: tok as string[] });
      } else if (typeof tok === "string") {
        rest.push({ kind: "single", value: tok });
      } else {
        return { rules, error: `规则 #${i} pattern token 非法` };
      }
    }
    rules.push({
      first,
      rest,
      decision,
      justification: typeof r.justification === "string" ? r.justification : undefined,
    });
  }
  return { rules };
}

/** 把 electron-store 中的用户规则与可写根应用到 sandbox 单例 */
export function applySandboxConfig(): void {
  const store = getSettingsStore();
  const userRules = (store.get("sandboxUserRules") as PrefixRule[] | undefined) ?? [];
  const roots = (store.get("sandboxExtraWritableRoots") as string[] | undefined) ?? [];
  setUserRules(userRules ?? []);
  setExtraWritableRoots(roots ?? []);
}

function getExcelBridgeForIpc(): ExcelConnectionBridge {
  return excelBridgeRef() ?? getOrCreateExcelBridge();
}
