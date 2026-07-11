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
  const installRoot = app.isPackaged ? path.dirname(process.execPath) : process.cwd();
  return path.join(installRoot, DATA_DIR_NAME);
}

export function getUserWritableDataPath(): string {
  return path.join(app.getPath("userData"), DATA_DIR_NAME);
}

export function getConfiguredDataPath(): string {
  const configured = bootstrapStore.get("dataPath") as string | undefined;
  return typeof configured === "string" && configured.trim() ? configured : "";
}

export function setConfiguredDataPath(dataPath: string): void {
  bootstrapStore.set("dataPath", dataPath);
}

export function normalizePathForCompare(targetPath: string): string {
  return path
    .resolve(targetPath)
    .replace(/[\\/]+$/, "")
    .toLowerCase();
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
