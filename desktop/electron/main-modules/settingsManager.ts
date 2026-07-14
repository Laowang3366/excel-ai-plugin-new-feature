/**
 * 设置管理器 — 持久化配置的读写
 *
 * 从 main.ts 提取，管理 electron-store 实例的创建、数据路径迁移、
 * AI 配置获取、窗口主题应用等。
 */

import * as path from "path";
import * as fs from "fs";
import { safeStorage } from "electron";
import Store from "electron-store";
import { AgentLoop } from "../agent/core/agentLoop";
import { AgentGraphStore } from "../agent/memory/agentGraphStore";
import { SessionStore } from "../agent/memory/sessionStore";
import { StateRuntimeStore } from "../agent/memory/stateRuntimeStore";
import type { AIClientConfig } from "../agent/providers/aiClient";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import { reloadKnowledgeRuntime, resetKnowledgeRuntime } from "../agent/runtime/knowledgeRuntime";
import { configureLogDirectory } from "../shared/logger";
import {
  copyDirectoryContents,
  getActiveDataPath,
  isPathInside,
  normalizePathForCompare,
  pathExists,
  setConfiguredDataPath,
  SETTINGS_STORE_NAME,
} from "./settingsDataPath";
import { AsyncResource } from "./asyncResource";
import {
  decryptProviderForRuntime,
  decryptSettingValueForRuntime,
  migrateSettingsSecrets,
  protectSettingValueForStorage,
  sanitizeSettingsForRenderer,
  type SettingsSecretCipher,
} from "./settingsSecrets";

export { getActiveDataPath };

const DEFAULT_SETTINGS = {
  aiProviders: {},
  activeProvider: "",
  permissionMode: "normal",
  showReasoning: true,
  language: "zh-CN",
  theme: "light",
  closeToTray: false,
  officeAutoCompactEnabled: false,
  dynamicArrayFunctionsEnabled: true,
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
};

const MIN_WINDOW_OPACITY = 0.55;
const MAX_WINDOW_OPACITY = 1;

export function normalizeWindowOpacity(value: unknown): number {
  const numericValue = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numericValue)) return MAX_WINDOW_OPACITY;
  const clamped = Math.min(Math.max(numericValue, MIN_WINDOW_OPACITY), MAX_WINDOW_OPACITY);
  return Math.round(clamped * 100) / 100;
}

let settingsStore = new Store(getSettingsStoreOptions(getActiveDataPath()));

const settingsSecretCipher: SettingsSecretCipher = {
  isAvailable: () => safeStorage?.isEncryptionAvailable() === true,
  encrypt: (value) => safeStorage.encryptString(value).toString("base64"),
  decrypt: (value) => safeStorage.decryptString(Buffer.from(value, "base64")),
};

export function getSettingsStore(): Store<typeof DEFAULT_SETTINGS> {
  return settingsStore;
}

export function initializeSettingsSecrets(): void {
  const current = settingsStore.store as Record<string, unknown>;
  const migrated = migrateSettingsSecrets(current, settingsSecretCipher);
  if (JSON.stringify(migrated) !== JSON.stringify(current)) {
    settingsStore.store = migrated as typeof DEFAULT_SETTINGS;
  }
}

export function getSettingsForRenderer(): Record<string, unknown> {
  return sanitizeSettingsForRenderer(settingsStore.store as Record<string, unknown>);
}

export function getSettingForRenderer(key: string): unknown {
  return getSettingsForRenderer()[key];
}

export function setSettingFromRenderer(key: string, value: unknown): unknown {
  const protectedValue = protectSettingValueForStorage(
    key,
    value,
    settingsStore.get(key as keyof typeof DEFAULT_SETTINGS),
    settingsSecretCipher
  );
  settingsStore.set(key as keyof typeof DEFAULT_SETTINGS, protectedValue as never);
  return getSettingForRenderer(key);
}

export function getRuntimeSettingValue(key: string): unknown {
  return decryptSettingValueForRuntime(
    key,
    settingsStore.get(key as keyof typeof DEFAULT_SETTINGS),
    settingsSecretCipher
  );
}

export function getProviderApiKey(providerId: string): string {
  const providers = (settingsStore.get("aiProviders") as Record<string, Record<string, unknown>>) || {};
  const provider = providers[providerId];
  if (!provider) return "";
  return String(decryptProviderForRuntime(provider, settingsSecretCipher).apiKey || "");
}

function getSettingsStoreOptions(dataPath?: string) {
  return dataPath
    ? {
        name: SETTINGS_STORE_NAME,
        cwd: path.join(dataPath, "settings"),
        defaults: DEFAULT_SETTINGS,
      }
    : { name: SETTINGS_STORE_NAME, defaults: DEFAULT_SETTINGS };
}

let sessionStore: SessionStore | null = null;
let agentGraphStore: AgentGraphStore | null = null;
const stateRuntimeResource = new AsyncResource(
  async () => {
    const sessionsRoot = path.join(getActiveDataPath(), "sessions");
    const store = new StateRuntimeStore(path.join(sessionsRoot, "state-runtime"));
    await store.init();
    return store;
  },
  (store) => store.close(),
);

/** 获取 AgentLoop 实例的回调（由 main.ts 注册），用于迁移后刷新 SessionStore */
let agentLoopsGetter: (() => AgentLoop[]) | null = null;

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
  if (migrationInProgress) {
    throw new Error("数据存储正在迁移，请稍后重试");
  }
  return openStateRuntimeStoreInstance();
}

export async function closeStateRuntimeStore(): Promise<void> {
  return stateRuntimeResource.close();
}

async function openStateRuntimeStoreInstance(): Promise<StateRuntimeStore> {
  return stateRuntimeResource.get();
}

async function resetSessionStore(): Promise<void> {
  await closeStateRuntimeStore();
  sessionStore = null;
  agentGraphStore = null;

  const agents = agentLoopsGetter?.() ?? [];
  if (agents.length > 0) {
    const nextSessionStore = getSessionStoreInstance();
    const nextStateRuntimeStore = await openStateRuntimeStoreInstance();
    for (const agent of agents) {
      agent.updateSessionStore(nextSessionStore);
      agent.updateStateRuntimeStore(nextStateRuntimeStore);
    }
  }
}

/** 迁移互斥锁 — 防止并发调用 migrateDataPath */
let migrationInProgress = false;

export async function migrateDataPath(
  targetDataPath: string,
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
  const agents = agentLoopsGetter?.() ?? [];
  if (agents.some((agent) => agent.getIsRunning())) {
    migrationInProgress = false;
    return { success: false, error: "请等待当前会话执行完成或停止后再迁移数据" };
  }

  const previousSettingsStore = settingsStore;
  const previousSessionStore = getSessionStoreInstance();
  const previousAgentGraphStore = agentGraphStore;
  let switchedDataPath = false;
  const targetDataPathExisted = await pathExists(nextDataPath);
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
    previousSessionStore.suspendWrites("数据存储正在迁移，请稍后重试");
    await previousSessionStore.flushRolloutWrites();
    await closeStateRuntimeStore();
    resetKnowledgeRuntime();

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

    setConfiguredDataPath(nextDataPath);
    switchedDataPath = true;
    settingsStore = nextSettingsStore;
    configureLogDirectory(path.join(nextDataPath, "logs"));
    await resetSessionStore();
    const nextKnowledgeRuntime = await reloadKnowledgeRuntime(getActiveAIConfig(), nextDataPath);
    if (!nextKnowledgeRuntime.store) {
      throw new Error(nextKnowledgeRuntime.error || "新数据目录的知识库初始化失败");
    }

    return { success: true, dataPath: nextDataPath };
  } catch (error) {
    if (switchedDataPath) {
      await closeStateRuntimeStore().catch(() => {});
      setConfiguredDataPath(currentDataPath);
      settingsStore = previousSettingsStore;
      configureLogDirectory(path.join(currentDataPath, "logs"));
    }
    sessionStore = previousSessionStore;
    agentGraphStore = previousAgentGraphStore;
    previousSessionStore.resumeWrites();
    resetKnowledgeRuntime();

    try {
      const restoredStateRuntime = await openStateRuntimeStoreInstance();
      for (const agent of agents) {
        agent.updateSessionStore(previousSessionStore);
        agent.updateStateRuntimeStore(restoredStateRuntime);
      }
      await reloadKnowledgeRuntime(getActiveAIConfig(), currentDataPath);
    } catch {
      // 返回原始迁移错误；恢复失败会由后续存储访问继续暴露。
    }

    try {
      if (!targetDataPathExisted && (await pathExists(nextDataPath))) {
        await fs.promises.rm(nextDataPath, { recursive: true, force: true });
      } else if (!targetSessionsExisted && (await pathExists(nextSessionsRoot))) {
        await fs.promises.rm(nextSessionsRoot, { recursive: true, force: true });
      }
    } catch {
      /* ignore cleanup errors */
    }

    return {
      success: false,
      error: error instanceof Error ? error.message : "迁移数据失败",
    };
  } finally {
    previousSessionStore.resumeWrites();
    migrationInProgress = false;
  }
}

export function isDataMigrationInProgress(): boolean {
  return migrationInProgress;
}

export function getActiveAIConfig(): AIClientConfig {
  const activeProviderId = settingsStore.get("activeProvider") as string;
  const providers = (settingsStore.get("aiProviders") as Record<string, any>) || {};

  if (!activeProviderId || !providers[activeProviderId]) {
    return {
      provider: "openai",
      apiKey: "",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };
  }

  const p = decryptProviderForRuntime(providers[activeProviderId], settingsSecretCipher) as any;
  const activeModelConfig = p.modelConfigs?.find((m: any) => m.name === p.model);
  return {
    provider: p.provider,
    apiKey: p.apiKey || "",
    baseUrl: p.baseUrl || p.defaultBaseUrl || "",
    model: p.model || p.defaultModel || "",
    apiFormat: p.apiFormat,
    customHeaders: p.customHeaders,
    contextWindowSize:
      activeModelConfig?.contextWindowSize || p.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
    compHash: activeModelConfig?.compHash || p.compHash,
    reasoningMode: activeModelConfig?.reasoningMode || p.reasoningMode,
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
