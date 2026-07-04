/**
 * 设置管理器 — 持久化配置的读写
 *
 * 从 main.ts 提取，管理 electron-store 实例的创建、数据路径迁移、
 * AI 配置获取、窗口主题应用等。
 */

import { app } from "electron";
import * as path from "path";
import * as fs from "fs";
import Store from "electron-store";
import { AgentLoop } from "../agent/core/agentLoop";
import { AgentGraphStore } from "../agent/memory/agentGraphStore";
import { SessionStore } from "../agent/memory/sessionStore";
import { StateRuntimeStore } from "../agent/memory/stateRuntimeStore";
import type { AIClientConfig } from "../agent/providers/aiClient";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import { reloadKnowledgeRuntime } from "../agent/runtime/knowledgeRuntime";
import { configureLogDirectory } from "../shared/logger";

const SETTINGS_STORE_NAME = "excel-ai-settings";
const DATA_DIR_NAME = "data";

const bootstrapStore = new Store({
  name: "excel-ai-bootstrap",
});

const DEFAULT_SETTINGS = {
  aiProviders: {},
  activeProvider: "",
  permissionMode: "normal",
  showReasoning: true,
  language: "zh-CN",
  theme: "light",
  closeToTray: false,
  officeAutoCompactEnabled: false,
  windowOpacity: 1,
  dataStoragePath: "",
  mineruApiToken: "",
  compactionConfig: {
    enabled: true,
    autoCompactThresholdPercent: 80,
    retainedUserMessageMaxTokens: 20000,
    summaryRetryCount: 1,
    midTurnThresholdRatio: 0.9,
  },
  // 沙箱命令策略 — 用户自定义规则与可写根目录
  // 见 docs/sandbox-implementation-plan.md
  sandboxUserRules: [] as unknown[],
  sandboxExtraWritableRoots: [] as string[],
};

const MIN_WINDOW_OPACITY = 0.55;
const MAX_WINDOW_OPACITY = 1;

export function normalizeWindowOpacity(value: unknown): number {
  const numericValue = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numericValue)) return MAX_WINDOW_OPACITY;
  const clamped = Math.min(Math.max(numericValue, MIN_WINDOW_OPACITY), MAX_WINDOW_OPACITY);
  return Math.round(clamped * 100) / 100;
}

migrateLegacyDefaultDataPathIfNeeded();

let settingsStore = new Store(getSettingsStoreOptions(getActiveDataPath()));

export function getSettingsStore(): Store<typeof DEFAULT_SETTINGS> {
  return settingsStore;
}

function getInstallDataPath(): string {
  const installRoot = app.isPackaged
    ? path.dirname(process.execPath)
    : process.cwd();
  return path.join(installRoot, DATA_DIR_NAME);
}

function getUserWritableDataPath(): string {
  return path.join(app.getPath("userData"), DATA_DIR_NAME);
}

function getLegacyRoamingDataPath(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant");
}

function getConfiguredDataPath(): string {
  const configured = bootstrapStore.get("dataPath") as string | undefined;
  return typeof configured === "string" && configured.trim() ? configured : "";
}

export function getActiveDataPath(): string {
  const configured = getConfiguredDataPath();
  if (configured) {
    if (ensureWritableDataPathSync(configured)) return configured;
    console.warn("配置的数据目录不可写，已回退到用户数据目录:", configured);
  }

  const installDataPath = getInstallDataPath();
  if (ensureWritableDataPathSync(installDataPath)) return installDataPath;

  const userWritableDataPath = getUserWritableDataPath();
  ensureWritableDataPathSync(userWritableDataPath);
  return userWritableDataPath;
}

function getSettingsStoreOptions(dataPath?: string) {
  return dataPath
    ? { name: SETTINGS_STORE_NAME, cwd: path.join(dataPath, "settings"), defaults: DEFAULT_SETTINGS }
    : { name: SETTINGS_STORE_NAME, defaults: DEFAULT_SETTINGS };
}

function normalizePathForCompare(targetPath: string): string {
  return path.resolve(targetPath).replace(/[\\/]+$/, "").toLowerCase();
}

function isPathInside(parentPath: string, childPath: string): boolean {
  const parent = normalizePathForCompare(parentPath);
  const child = normalizePathForCompare(childPath);
  return child === parent || child.startsWith(`${parent}${path.sep.toLowerCase()}`);
}

function ensureWritableDataPathSync(dataPath: string): boolean {
  try {
    const settingsDir = path.join(dataPath, "settings");
    fs.mkdirSync(settingsDir, { recursive: true });
    const probePath = path.join(settingsDir, `.write-test-${process.pid}-${Date.now()}`);
    fs.writeFileSync(probePath, "ok", "utf8");
    fs.rmSync(probePath, { force: true });
    return true;
  } catch {
    return false;
  }
}

function migrateLegacyDefaultDataPathIfNeeded(): void {
  if (getConfiguredDataPath()) return;

  const legacyDataPath = getLegacyRoamingDataPath();
  const installDataPath = getInstallDataPath();
  const nextDataPath = ensureWritableDataPathSync(installDataPath)
    ? installDataPath
    : getUserWritableDataPath();
  if (normalizePathForCompare(legacyDataPath) === normalizePathForCompare(nextDataPath)) return;
  if (!fs.existsSync(legacyDataPath)) return;

  try {
    fs.mkdirSync(nextDataPath, { recursive: true });

    const legacySettingsPath = path.join(legacyDataPath, `${SETTINGS_STORE_NAME}.json`);
    const nextSettingsDir = path.join(nextDataPath, "settings");
    const nextSettingsPath = path.join(nextSettingsDir, `${SETTINGS_STORE_NAME}.json`);
    if (fs.existsSync(legacySettingsPath) && !fs.existsSync(nextSettingsPath)) {
      fs.mkdirSync(nextSettingsDir, { recursive: true });
      fs.copyFileSync(legacySettingsPath, nextSettingsPath);
    }

    const legacySessionsRoot = path.join(legacyDataPath, "sessions");
    const nextSessionsRoot = path.join(nextDataPath, "sessions");
    if (fs.existsSync(legacySessionsRoot) && !fs.existsSync(nextSessionsRoot)) {
      copyDirectoryContentsSync(legacySessionsRoot, nextSessionsRoot);
    }

    const legacyKnowledgeRoot = path.join(legacyDataPath, "knowledge");
    const nextKnowledgeRoot = path.join(nextDataPath, "knowledge");
    if (fs.existsSync(legacyKnowledgeRoot) && !fs.existsSync(nextKnowledgeRoot)) {
      copyDirectoryContentsSync(legacyKnowledgeRoot, nextKnowledgeRoot);
    }

    const legacyLogsRoot = path.join(legacyDataPath, "logs");
    const nextLogsRoot = path.join(nextDataPath, "logs");
    if (fs.existsSync(legacyLogsRoot) && !fs.existsSync(nextLogsRoot)) {
      copyDirectoryContentsSync(legacyLogsRoot, nextLogsRoot);
    }
  } catch (error) {
    console.warn("迁移默认数据目录失败:", error);
  }
}

function copyDirectoryContentsSync(sourceDir: string, targetDir: string): void {
  if (!fs.existsSync(sourceDir)) return;
  fs.mkdirSync(targetDir, { recursive: true });
  const entries = fs.readdirSync(sourceDir, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = path.join(targetDir, entry.name);
    if (entry.isDirectory()) {
      copyDirectoryContentsSync(sourcePath, targetPath);
      continue;
    }
    if (entry.isFile() && !fs.existsSync(targetPath)) {
      fs.copyFileSync(sourcePath, targetPath);
    }
  }
}

async function pathExists(targetPath: string): Promise<boolean> {
  try {
    await fs.promises.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function copyDirectoryContents(sourceDir: string, targetDir: string): Promise<void> {
  if (!(await pathExists(sourceDir))) return;
  await fs.promises.mkdir(targetDir, { recursive: true });
  const entries = await fs.promises.readdir(sourceDir, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = path.join(targetDir, entry.name);
    if (entry.isDirectory()) {
      await copyDirectoryContents(sourcePath, targetPath);
      continue;
    }
    if (entry.isFile() && !(await pathExists(targetPath))) {
      await fs.promises.copyFile(sourcePath, targetPath);
    }
  }
}

let sessionStore: SessionStore | null = null;
let agentGraphStore: AgentGraphStore | null = null;
let stateRuntimeStore: StateRuntimeStore | null = null;

/** 获取 AgentLoop 实例的回调（由 main.ts 注册），用于迁移后刷新 SessionStore */
let agentLoopGetter: (() => AgentLoop | null) | null = null;
let agentLoopsGetter: (() => AgentLoop[]) | null = null;

export function setAgentLoopGetter(fn: () => AgentLoop | null): void {
  agentLoopGetter = fn;
}

export function setAgentLoopsGetter(fn: () => AgentLoop[]): void {
  agentLoopsGetter = fn;
}

export function getSessionStoreInstance(): SessionStore {
  if (!sessionStore) {
    const sessionsRoot = path.join(getActiveDataPath(), "sessions");
    sessionStore = new SessionStore(sessionsRoot);
  }
  return sessionStore;
}

export function getAgentGraphStoreInstance(): AgentGraphStore {
  if (!agentGraphStore) {
    const sessionsRoot = path.join(getActiveDataPath(), "sessions");
    agentGraphStore = new AgentGraphStore(sessionsRoot);
  }
  return agentGraphStore;
}

export async function getStateRuntimeStoreInstance(): Promise<StateRuntimeStore> {
  if (!stateRuntimeStore) {
    const sessionsRoot = path.join(getActiveDataPath(), "sessions");
    stateRuntimeStore = new StateRuntimeStore(path.join(sessionsRoot, "state-runtime"));
    await stateRuntimeStore.init();
  }
  return stateRuntimeStore;
}

export async function closeStateRuntimeStore(): Promise<void> {
  if (!stateRuntimeStore) return;
  await stateRuntimeStore.close();
  stateRuntimeStore = null;
}

export function resetSessionStore(): void {
  const previousStateRuntimeStore = stateRuntimeStore;
  sessionStore = null;
  agentGraphStore = null;
  stateRuntimeStore = null;
  void previousStateRuntimeStore?.close().catch(() => {});
  // 数据迁移后，刷新 AgentLoop 的 SessionStore 引用，防止指向旧路径
  const fallbackAgent = agentLoopGetter?.();
  const agents = agentLoopsGetter?.() ?? (fallbackAgent ? [fallbackAgent] : []);
  if (agents.length > 0) {
    const nextSessionStore = getSessionStoreInstance();
    for (const agent of agents) {
      agent.updateSessionStore(nextSessionStore);
    }
    void getStateRuntimeStoreInstance()
      .then((store) => {
        for (const agent of agents) {
          agent.updateStateRuntimeStore(store);
        }
      })
      .catch(() => {});
  }
}

/** 迁移互斥锁 — 防止并发调用 migrateDataPath */
let migrationInProgress = false;

export async function migrateDataPath(
  targetDataPath: string
): Promise<{ success: boolean; dataPath?: string; error?: string }> {
  // 互斥锁
  if (migrationInProgress) {
    return { success: false, error: "数据迁移正在进行中，请稍后重试" };
  }
  migrationInProgress = true;

  const trimmedPath = targetDataPath.trim();
  if (!trimmedPath) {
    migrationInProgress = false;
    return { success: false, error: "请选择有效的数据存储目录" };
  }

  const nextDataPath = path.resolve(trimmedPath);
  const currentDataPath = getActiveDataPath();
  const currentSessionsRoot = path.join(currentDataPath, "sessions");
  const nextSessionsRoot = path.join(nextDataPath, "sessions");
  const currentKnowledgeRoot = path.join(currentDataPath, "knowledge");
  const nextKnowledgeRoot = path.join(nextDataPath, "knowledge");
  const currentLogsRoot = path.join(currentDataPath, "logs");
  const nextLogsRoot = path.join(nextDataPath, "logs");
  const targetSessionsExisted = await pathExists(nextSessionsRoot);
  const sourceSessionEntries = await fs.promises.readdir(currentSessionsRoot).catch(() => []);

  if (normalizePathForCompare(currentDataPath) === normalizePathForCompare(nextDataPath)) {
    migrationInProgress = false;
    return { success: true, dataPath: currentDataPath };
  }

  if (isPathInside(currentSessionsRoot, nextSessionsRoot)) {
    migrationInProgress = false;
    return { success: false, error: "新目录不能位于当前会话数据目录内部" };
  }

  try {
    await fs.promises.mkdir(nextDataPath, { recursive: true });
    await copyDirectoryContents(currentSessionsRoot, nextSessionsRoot);
    await copyDirectoryContents(currentKnowledgeRoot, nextKnowledgeRoot);
    await copyDirectoryContents(currentLogsRoot, nextLogsRoot);

    // 迁移后验证：确认目标目录已包含数据
    if (sourceSessionEntries.length > 0) {
      const targetEntries = await fs.promises.readdir(nextSessionsRoot).catch(() => []);
      if (targetEntries.length === 0) {
        throw new Error("迁移验证失败：目标会话目录为空，可能复制失败");
      }
    }

    const currentSettings = settingsStore.store;
    const nextSettingsStore = new Store(getSettingsStoreOptions(nextDataPath));
    nextSettingsStore.store = {
      ...DEFAULT_SETTINGS,
      ...currentSettings,
      dataStoragePath: nextDataPath,
    };

    bootstrapStore.set("dataPath", nextDataPath);
    settingsStore = nextSettingsStore;
    configureLogDirectory(path.join(nextDataPath, "logs"));
    // resetSessionStore 内部会自动刷新 AgentLoop 的 SessionStore 引用
    resetSessionStore();
    await reloadKnowledgeRuntime(getActiveAIConfig(), nextDataPath);

    return { success: true, dataPath: nextDataPath };
  } catch (error) {
    // 回滚：清理不完整的目标目录（不删除 bootstrapStore 未切换，源数据安全）
    try {
      if (!targetSessionsExisted && await pathExists(nextSessionsRoot)) {
        await fs.promises.rm(nextSessionsRoot, { recursive: true, force: true });
      }
    } catch { /* ignore cleanup errors */ }

    return {
      success: false,
      error: error instanceof Error ? error.message : "迁移数据失败",
    };
  } finally {
    migrationInProgress = false;
  }
}

export function getActiveAIConfig(): AIClientConfig {
  const activeProviderId = settingsStore.get("activeProvider") as string;
  const providers = settingsStore.get("aiProviders") as Record<string, any> || {};

  if (!activeProviderId || !providers[activeProviderId]) {
    return {
      provider: "openai",
      apiKey: "",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };
  }

  const p = providers[activeProviderId];
  const activeModelConfig = p.modelConfigs?.find((m: any) => m.name === p.model);
  return {
    provider: p.provider,
    apiKey: p.apiKey || "",
    baseUrl: p.baseUrl || p.defaultBaseUrl || "",
    model: p.model || p.defaultModel || "",
    apiFormat: p.apiFormat,
    customHeaders: p.customHeaders,
    enableReasoning: p.enableReasoning || false,
    contextWindowSize: activeModelConfig?.contextWindowSize || p.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
    compHash: activeModelConfig?.compHash || p.compHash,
    reasoningMode: activeModelConfig?.reasoningMode || p.reasoningMode || (p.enableReasoning ? "high" : undefined),
  };
}

/** 应用窗口主题 */
export function applyWindowTheme(mainWindow: Electron.BrowserWindow | null): void {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  const theme = settingsStore.get("theme") === "dark" ? "dark" : "light";
  mainWindow.setBackgroundColor(theme === "dark" ? "#0f172a" : "#ffffff");
  mainWindow.setTitleBarOverlay({
    color: theme === "dark" ? "#0b1220" : "#eef5fb",
    symbolColor: theme === "dark" ? "#f8fafc" : "#111827",
    height: 36,
  });
}

/** Apply the configured main window opacity. */
export function applyWindowOpacity(mainWindow: Electron.BrowserWindow | null): void {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  mainWindow.setOpacity(normalizeWindowOpacity(settingsStore.get("windowOpacity")));
}
