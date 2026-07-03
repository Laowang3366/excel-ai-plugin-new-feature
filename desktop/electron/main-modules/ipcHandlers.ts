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
import { refreshKnowledgeRuntime, type AgentLoopManager } from "../agent/runtime/agentRuntime";
import { createAIClient } from "../agent/providers/aiClient";
import { getOrCreateExcelBridge } from "../agent/runtime/bridgeRegistry";
import type { ExcelConnectionBridge } from "../agent/tools/contracts/excel";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import { buildCompactionConfig, type SavedCompactionConfig } from "../agent/runtime/compactionRuntime";
import { registerAgentIpcHandlers } from "../agent/interaction/ipcAgentHandlers";
import {
  getSettingsStore,
  getActiveDataPath,
  getActiveAIConfig,
  getSessionStoreInstance,
  getStateRuntimeStoreInstance,
  getAgentGraphStoreInstance,
  migrateDataPath,
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
  AiListModelsInput,
  AiTestConnectionInput,
  MigrateDataPathInput,
} from "../shared/ipcSchemas";
import { createLogger } from "../shared/logger";
import { parseFilesWithMineru, parseFilesWithMineruAgent, type MineruParsedDocument } from "./mineruOcr";
import { parseFilesLocally } from "../agent/tools/executors/localDocumentParser";
import {
  buildInvoiceFieldFallback,
  buildRowsFromFields,
  isLikelyInvoiceText,
  mergeInvoiceFields,
} from "./invoiceFieldExtraction";
import {
  normalizeOcrMode,
  isLikelyInvoiceFileList,
  normalizePlainOcrText,
  type OcrInvoiceItem,
  type OcrVisionResult,
} from "./ocrModeDetection";
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
  registerAgentIpcHandlers({
    mainWindowRef,
    agentLoopRef,
    agentLoopManagerRef,
    getSessionStoreInstance,
    getStateRuntimeStoreInstance,
    getAgentGraphStoreInstance,
  });

  // ---- 应用信息 ----
  registerOcrIpcHandler();

  ipcMain.handle("app:getDataPath", () => getActiveDataPath());

  ipcMain.handle("app:openPath", async (_event, targetPath: string) => {
    return await shell.openPath(targetPath);
  });

  ipcMain.handle("app:openExternal", async (_event, targetUrl: string) => {
    try {
      const url = new URL(targetUrl);
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

  ipcMain.handle("window:setAlwaysOnTop", (_event, enabled: boolean) => {
    mainWindowRef()?.setAlwaysOnTop(enabled);
    return mainWindowRef()?.isAlwaysOnTop() ?? false;
  });

  ipcMain.handle("window:getDisplayMode", () => {
    return getWindowDisplayMode();
  });

  ipcMain.handle("window:setDisplayMode", (_event, mode: WindowDisplayMode) => {
    if (mode !== "normal" && mode !== "compact") {
      return getWindowDisplayMode();
    }
    return setWindowDisplayMode(mainWindowRef(), mode);
  });

  // ---- 设置相关 ----
  ipcMain.handle("settings:get", (_event, key: string) => {
    return getSettingsStore().get(key);
  });

  ipcMain.handle("settings:set", async (_event, key: string, value: unknown) => {
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

  ipcMain.handle("excel:selectHost", async (_event, host: "excel" | "wps") => {
    const bridge = excelBridgeRef();
    if (!bridge) return { connected: false, host: "unknown", error: "Bridge not available" };
    return await bridge.selectHost(host);
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

  ipcMain.handle("excel:readRange", async (_event, sheetName: string, range: string, expand?: "none" | "spill" | "currentArray" | "currentRegion") => {
    const bridge = excelBridgeRef();
    if (!bridge) return { values: [[]] };
    try { return await bridge.readRange(sheetName, range, expand); }
    catch { return { values: [[]] }; }
  });

  ipcMain.handle("excel:inspectWorkbook", async () => {
    const bridge = excelBridgeRef();
    if (!bridge) return null;
    try { return await bridge.inspectWorkbook(); }
    catch { return null; }
  });

  ipcMain.handle("excel:writeRange", async (_event, sheetName: string, range: string, values: unknown[][]) => {
    const bridge = excelBridgeRef();
    if (!bridge) return { success: false, error: "Excel 未连接" };
    try {
      await bridge.writeRange(sheetName, range, values);
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
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("dialog:openFolder", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      title: "选择文件夹",
      properties: ["openDirectory"],
    });
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  // ---- 文件夹文件列表 ----
  ipcMain.handle("folder:listFiles", async (_event, folderPath: string) => {
    try {
      const officeExts = new Set([".xlsx", ".xls", ".csv", ".doc", ".docx", ".ppt", ".pptx"]);
      const entries = await fs.promises.readdir(folderPath, { withFileTypes: true });
      const files = entries
        .filter((e) => e.isFile() && officeExts.has(path.extname(e.name).toLowerCase()))
        .map((e) => {
          const fullPath = path.join(folderPath, e.name);
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
  ipcMain.handle("file:readAsBase64", async (_event, filePath: string) => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      const ext = path.extname(filePath).toLowerCase().replace(".", "");
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
        fileName: path.basename(filePath),
        size: buffer.length,
      };
    } catch (err: any) {
      return { error: err.message };
    }
  });

  // ---- 临时文件操作（截图粘贴等） ----
  ipcMain.handle("file:writeTempFile", async (_event, data: unknown) => {
    try {
      const input = data as { prefix?: string; suffix?: string; data: string };
      if (!input.data || typeof input.data !== "string") {
        return { success: false, error: "缺少 data 参数" };
      }
      const prefix = input.prefix?.replace(/[^a-zA-Z0-9_-]/g, "") || "clipboard";
      const suffix = input.suffix?.replace(/[^a-zA-Z0-9.]/g, "") || ".png";
      const tmpDir = await import("os").then((os) => os.tmpdir());
      const fileName = `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}${suffix}`;
      const filePath = path.join(tmpDir, fileName);
      const buffer = Buffer.from(input.data, "base64");
      await fs.promises.writeFile(filePath, buffer);
      return { success: true, filePath };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  // ---- 文件操作（回收站/打开/复制路径/显示） ----
  ipcMain.handle("file:trashFile", async (_event, filePath: string) => {
    try {
      await shell.trashItem(filePath);
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:openFile", async (_event, filePath: string) => {
    try {
      const result = await shell.openPath(filePath);
      if (result) return { success: false, error: result };
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:copyPath", (_event, filePath: string) => {
    try {
      clipboard.writeText(filePath);
      return { success: true };
    } catch (err: any) {
      return { success: false };
    }
  });

  ipcMain.handle("file:revealInExplorer", (_event, filePath: string) => {
    try {
      shell.showItemInFolder(filePath);
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
    if (!Array.isArray(rules)) {
      return { success: false, error: "rules 必须为数组" };
    }
    const normalized = normalizeUserRules(rules);
    if (normalized.error) {
      return { success: false, error: normalized.error };
    }
    getSettingsStore().set("sandboxUserRules", normalized.rules);
    applySandboxConfig();
    return { success: true };
  });

  ipcMain.handle("sandbox:setWritableRoots", (_event, roots: unknown) => {
    if (!Array.isArray(roots)) {
      return { success: false, error: "roots 必须为数组" };
    }
    const clean = roots.filter((r): r is string => typeof r === "string" && r.trim().length > 0);
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

function registerOcrIpcHandler(): void {
  try {
    ipcMain.removeHandler("ocr:recognize");
  } catch {
    // Handler may not exist on first registration.
  }

  ipcMain.handle("ocr:recognize", async (_event, mode: unknown, filePaths: unknown) => {
    try {
      return await recognizeWithOcrFallbacks(mode, filePaths);
    } catch (err: any) {
      return emptyOcrResult(normalizeOcrMode(mode), [err?.message || "OCR 识别失败"]);
    }
  });
}

async function recognizeWithOcrFallbacks(
  rawMode: unknown,
  rawFilePaths: unknown,
): Promise<OcrVisionResult> {
  const mode = normalizeOcrMode(rawMode);
  const filePaths = normalizeOcrFilePaths(rawFilePaths);
  const effectiveMode = mode === "invoice" || isLikelyInvoiceFileList(filePaths) ? "invoice" : "image";

  const parsed = await parseFilesWithOcrFallbacks(filePaths);
  if (!hasAnyUsefulParsedDocument(parsed.documents)) {
    return emptyOcrResult(effectiveMode, [
      "未提取到可用 OCR 文本或表格，无法抽取字段",
      ...parsed.errors,
      ...formatParsedDocumentErrors(parsed.documents),
    ]);
  }

  const result = effectiveMode === "invoice" || isLikelyInvoiceDocuments(parsed.documents)
    ? await extractInvoiceFieldsFromMineruDocuments(parsed.documents)
    : buildMineruOcrResult(parsed.documents);

  return {
    ...result,
    errors: [
      ...result.errors,
      ...formatParsedDocumentErrors(parsed.documents),
    ],
  };
}

async function parseFilesWithOcrFallbacks(
  filePaths: string[],
): Promise<{ documents: MineruParsedDocument[]; provider: "mineru" | "mineru-agent" | "local"; errors: string[] }> {
  const errors: string[] = [];
  const mineruToken = getConfiguredMineruToken();

  if (mineruToken) {
    try {
      const documents = await parseFilesWithMineru(filePaths, mineruToken);
      if (hasAnyUsefulParsedDocument(documents)) {
        return { documents, provider: "mineru", errors: [] };
      }
      errors.push(formatMineruDocumentErrors(documents) || "MinerU 标准解析未返回可用文本");
    } catch (error: any) {
      errors.push(`MinerU 标准解析失败：${error?.message || "未知错误"}`);
    }
  }

  try {
    const documents = await parseFilesWithMineruAgent(filePaths);
    if (hasAnyUsefulParsedDocument(documents)) {
      return { documents, provider: "mineru-agent", errors: [] };
    }
    errors.push(formatMineruDocumentErrors(documents) || "MinerU 免费解析未返回可用文本");
  } catch (error: any) {
    errors.push(`MinerU 免费解析失败：${error?.message || "未知错误"}`);
  }

  const localDocuments = await parseFilesLocally(filePaths);
  return {
    documents: localDocuments,
    provider: "local",
    errors: hasAnyUsefulParsedDocument(localDocuments) ? [] : errors,
  };
}

function hasAnyUsefulParsedDocument(documents: Array<{ text: string; rows: string[][] }>): boolean {
  return documents.some((document) => document.text.trim().length > 0 || document.rows.length > 0);
}

function formatParsedDocumentErrors(documents: Array<{ filename: string; error?: string }>): string[] {
  return documents
    .filter((document) => document.error && !/^local_ocr_unsupported|local_unsupported|local_empty$/.test(document.error))
    .map((document) => `${document.filename}: ${document.error}`);
}

function normalizeOcrFilePaths(rawFilePaths: unknown): string[] {
  if (!Array.isArray(rawFilePaths)) {
    throw new Error("OCR 文件列表必须是数组");
  }

  const filePaths = rawFilePaths
    .filter((value): value is string => typeof value === "string")
    .map((value) => value.trim())
    .filter(Boolean);
  if (filePaths.length === 0) {
    throw new Error("请先选择要识别的图片或 PDF");
  }

  return filePaths;
}

function isLikelyInvoiceDocuments(documents: MineruParsedDocument[]): boolean {
  return documents.some((document) =>
    isLikelyInvoiceText(`${document.filename}\n${document.text}`)
  );
}

function getConfiguredMineruToken(): string {
  const store = getSettingsStore();
  const configured = store.get("mineruApiToken") || store.get("ocrMineruApiToken");
  const tokenFromSettings = typeof configured === "string" ? configured.trim() : "";
  return tokenFromSettings || (process.env.MINERU_API_TOKEN || "").trim();
}

function buildMineruOcrResult(documents: MineruParsedDocument[]): OcrVisionResult {
  const text = buildCombinedMineruText(documents);
  return {
    kind: "image",
    text,
    rows: documents.flatMap((document) => document.rows),
    fields: {},
    invoices: [],
    errors: normalizeStringArray(formatMineruDocumentErrors(documents).split("\n")),
  };
}

async function extractInvoiceFieldsFromMineruDocuments(
  documents: MineruParsedDocument[],
): Promise<OcrVisionResult> {
  const fallback = buildMineruInvoiceFallbackResult(documents);
  try {
    const aiClient = createAIClient(getActiveAIConfig());
    const result = await aiClient.chat({
      messages: [{
        role: "user",
        content: buildMineruInvoicePrompt(documents),
      }],
      maxTokens: 4000,
      temperature: 0,
      reasoningMode: "off",
    });
    const normalized = normalizeOcrVisionResult("invoice", result.content || "");
    const extractedInvoices = normalizeExtractedInvoiceItems(normalized, documents);
    const mergedInvoices = mergeMineruInvoiceItems(extractedInvoices, documents);
    return {
      kind: "invoice",
      text: normalized.text || fallback.text,
      rows: normalized.rows.length > 0 ? normalized.rows : fallback.rows,
      fields: Object.keys(normalized.fields).length > 0 ? normalized.fields : fallback.fields,
      invoices: mergedInvoices.length > 0 ? mergedInvoices : fallback.invoices,
      errors: [
        ...fallback.errors,
        ...normalized.errors,
      ],
    };
  } catch (error: any) {
    return {
      ...fallback,
      errors: [
        ...fallback.errors,
        `发票字段抽取失败，已保留 MinerU OCR 文本：${error?.message || "未知错误"}`,
      ],
    };
  }
}

function normalizeExtractedInvoiceItems(
  normalized: OcrVisionResult,
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  if (normalized.invoices.length > 0) return normalized.invoices;
  if (Object.keys(normalized.fields).length === 0) return [];
  if (documents.length <= 1) {
    const document = documents[0];
    return [{
      filename: document?.filename || "识别文本",
      text: normalized.text || document?.text || "",
      fields: normalized.fields,
      rows: normalized.rows,
    }];
  }
  return [{
    filename: documents[0]?.filename || "识别文本",
    text: normalized.text || documents[0]?.text || "",
    fields: normalized.fields,
    rows: normalized.rows,
  }];
}

function buildMineruInvoiceFallbackResult(documents: MineruParsedDocument[]): OcrVisionResult {
  const extracted = buildInvoiceFieldFallback(documents);
  return {
    kind: "invoice",
    text: buildCombinedMineruText(documents),
    rows: extracted.rows.length > 0 ? extracted.rows : documents.flatMap((document) => document.rows),
    fields: extracted.fields,
    invoices: extracted.invoices,
    errors: normalizeStringArray(formatMineruDocumentErrors(documents).split("\n")),
  };
}

function buildMineruInvoicePrompt(documents: MineruParsedDocument[]): string {
  return [
    "下面是 MinerU 通用 OCR/版面解析得到的发票 Markdown 文本。",
    "请只基于这些文本抽取发票字段，不要编造缺失信息。",
    "只返回严格 JSON，不要 Markdown，不要解释。",
    "JSON 结构必须是：",
    "{\"kind\":\"invoice\",\"text\":\"合并后的可读文本\",\"rows\":[[\"列1\",\"列2\"]],\"fields\":{\"字段\":\"值\"},\"invoices\":[{\"filename\":\"文件名\",\"text\":\"文本\",\"fields\":{\"发票号码\":\"\",\"开票日期\":\"\",\"购买方名称\":\"\",\"销售方名称\":\"\",\"金额\":\"\",\"税额\":\"\",\"价税合计\":\"\"},\"rows\":[[\"列1\",\"列2\"]]}],\"errors\":[]}",
    "字段优先包含：发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、发票类型、校验码、备注。",
    "每个输入文件都要在 invoices 中返回一项；未识别字段填空字符串。",
    "",
    buildLimitedMineruSource(documents),
  ].join("\n");
}

function buildCombinedMineruText(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.text.trim())
    .map((document) => `## ${document.filename}\n${document.text.trim()}`)
    .join("\n\n");
}

function buildLimitedMineruSource(documents: MineruParsedDocument[]): string {
  const perDocumentLimit = 12_000;
  const totalLimit = 40_000;
  let used = 0;
  const sections: string[] = [];

  for (const document of documents) {
    if (!document.text.trim()) continue;
    const remaining = totalLimit - used;
    if (remaining <= 0) break;
    const clippedText = document.text.trim().slice(0, Math.min(perDocumentLimit, remaining));
    used += clippedText.length;
    sections.push(`### 文件：${document.filename}\n${clippedText}`);
  }

  return sections.join("\n\n");
}

function mergeMineruInvoiceItems(
  extractedInvoices: OcrInvoiceItem[],
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  const fallbackInvoices = buildInvoiceFieldFallback(documents).invoices;
  if (documents.length === 0) return extractedInvoices;
  return documents.map((document, index) => {
    const extracted = extractedInvoices.find((invoice) => invoice.filename === document.filename)
      || extractedInvoices[index];
    const fallback = fallbackInvoices.find((invoice) => invoice.filename === document.filename)
      || fallbackInvoices[index];
    const fields = mergeInvoiceFields(fallback?.fields, extracted?.fields);
    return {
      filename: extracted?.filename || document.filename,
      text: extracted?.text || document.text,
      fields,
      rows: extracted?.rows?.length
        ? extracted.rows
        : Object.keys(fields).length > 0
        ? buildRowsFromFields([fields])
        : fallback?.rows?.length
        ? fallback.rows
        : document.rows,
      error: document.error || extracted?.error,
    };
  });
}

function formatMineruDocumentErrors(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.error)
    .map((document) => `${document.filename}: ${document.error}`)
    .join("\n");
}

export function normalizeOcrVisionResult(mode: "image" | "invoice", content: string): OcrVisionResult {
  const parsed = parseJsonObject(content);
  if (!parsed) {
    return normalizePlainOcrText(mode, content);
  }

  const result = parsed as Record<string, unknown>;
  const invoices = normalizeInvoices(result.invoices);
  return {
    kind: mode,
    text: typeof result.text === "string" ? result.text : invoices.map((item) => item.text).filter(Boolean).join("\n"),
    rows: normalizeRows(result.rows),
    fields: normalizeFields(result.fields),
    invoices,
    errors: normalizeStringArray(result.errors),
  };
}

function parseJsonObject(content: string): Record<string, unknown> | null {
  const trimmed = content.trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  const candidate = fenced ? fenced[1].trim() : trimmed;
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  try {
    const parsed = JSON.parse(candidate.slice(start, end + 1));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : null;
  } catch {
    return null;
  }
}

function emptyOcrResult(mode: "image" | "invoice", errors: string[] = []): OcrVisionResult {
  return {
    kind: mode,
    text: "",
    rows: [],
    fields: {},
    invoices: [],
    errors,
  };
}

function normalizeInvoices(value: unknown): OcrInvoiceItem[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((item): OcrInvoiceItem[] => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return [];
    const raw = item as Record<string, unknown>;
    return [{
      filename: typeof raw.filename === "string" ? raw.filename : "",
      text: typeof raw.text === "string" ? raw.text : "",
      fields: normalizeFields(raw.fields),
      rows: normalizeRows(raw.rows),
      error: typeof raw.error === "string" ? raw.error : undefined,
    }];
  });
}

function normalizeFields(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const fields: Record<string, string> = {};
  for (const [key, rawValue] of Object.entries(value)) {
    if (!key.trim()) continue;
    fields[key] = rawValue === null || rawValue === undefined
      ? ""
      : typeof rawValue === "string"
      ? rawValue
      : String(rawValue);
  }
  return fields;
}

function normalizeRows(value: unknown): string[][] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((row): string[][] => {
    if (!Array.isArray(row)) return [];
    return [row.map((cell) => cell === null || cell === undefined ? "" : String(cell))];
  });
}

function normalizeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => item.trim())
    .filter(Boolean);
}
