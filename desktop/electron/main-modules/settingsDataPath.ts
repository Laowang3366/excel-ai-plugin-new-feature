import { app } from "electron";
import * as fs from "fs";
import * as path from "path";
import Store from "electron-store";
import { createLogger } from "../shared/logger";

export const SETTINGS_STORE_NAME = "excel-ai-settings";

const DATA_DIR_NAME = "data";
const dataPathLogger = createLogger("SettingsDataPath");
const bootstrapStore = new Store({
  name: "excel-ai-bootstrap",
});

export function getInstallDataPath(): string {
  const installRoot = app.isPackaged
    ? path.dirname(process.execPath)
    : process.cwd();
  return path.join(installRoot, DATA_DIR_NAME);
}

export function getUserWritableDataPath(): string {
  return path.join(app.getPath("userData"), DATA_DIR_NAME);
}

function getLegacyRoamingDataPath(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant");
}

export function getConfiguredDataPath(): string {
  const configured = bootstrapStore.get("dataPath") as string | undefined;
  return typeof configured === "string" && configured.trim() ? configured : "";
}

export function setConfiguredDataPath(dataPath: string): void {
  bootstrapStore.set("dataPath", dataPath);
}

export function normalizePathForCompare(targetPath: string): string {
  return path.resolve(targetPath).replace(/[\\/]+$/, "").toLowerCase();
}

export function isPathInside(parentPath: string, childPath: string): boolean {
  const parent = normalizePathForCompare(parentPath);
  const child = normalizePathForCompare(childPath);
  return child === parent || child.startsWith(`${parent}${path.sep.toLowerCase()}`);
}

export function ensureWritableDataPathSync(dataPath: string): boolean {
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

export function getActiveDataPath(): string {
  const configured = getConfiguredDataPath();
  if (configured) {
    if (ensureWritableDataPathSync(configured)) return configured;
    dataPathLogger.warn("配置的数据目录不可写，已回退到用户数据目录", { configured });
  }

  const installDataPath = getInstallDataPath();
  if (ensureWritableDataPathSync(installDataPath)) return installDataPath;

  const userWritableDataPath = getUserWritableDataPath();
  ensureWritableDataPathSync(userWritableDataPath);
  return userWritableDataPath;
}

export function migrateLegacyDefaultDataPathIfNeeded(): void {
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
    dataPathLogger.warn(
      "迁移默认数据目录失败",
      error instanceof Error ? { message: error.message, stack: error.stack } : { error: String(error) }
    );
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

export async function pathExists(targetPath: string): Promise<boolean> {
  try {
    await fs.promises.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

export async function copyDirectoryContents(sourceDir: string, targetDir: string): Promise<void> {
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
