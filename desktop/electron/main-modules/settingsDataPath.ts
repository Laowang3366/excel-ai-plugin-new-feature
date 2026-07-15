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

  const userWritableDataPath = getUserWritableDataPath();
  const installDataPath = getInstallDataPath();
  if (
    normalizePathForCompare(installDataPath) !== normalizePathForCompare(userWritableDataPath) &&
    hasMeaningfulDataSync(installDataPath) &&
    !hasMeaningfulDataSync(userWritableDataPath)
  ) {
    try {
      if (containsSymbolicLinkSync(installDataPath)) {
        throw new Error("旧数据目录包含符号链接或联接");
      }
      migrateLegacyDataDirectorySync(installDataPath, userWritableDataPath);
      setConfiguredDataPath(userWritableDataPath);
      dataPathLogger.info("已将旧安装目录数据迁移到用户隔离目录", {
        from: installDataPath,
        to: userWritableDataPath,
      });
    } catch (error) {
      dataPathLogger.warn("旧安装目录数据自动迁移失败", {
        from: installDataPath,
        to: userWritableDataPath,
        error: error instanceof Error ? error.message : String(error),
      });
      if (ensureWritableDataPathSync(installDataPath)) return installDataPath;
    }
  }

  ensureWritableDataPathSync(userWritableDataPath);
  return userWritableDataPath;
}

function hasMeaningfulDataSync(dataPath: string): boolean {
  try {
    return fs.readdirSync(dataPath).some((entry) => entry !== ".DS_Store");
  } catch {
    return false;
  }
}

function containsSymbolicLinkSync(directory: string): boolean {
  if (fs.lstatSync(directory).isSymbolicLink()) return true;
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (entry.isSymbolicLink()) return true;
    if (entry.isDirectory() && containsSymbolicLinkSync(path.join(directory, entry.name))) {
      return true;
    }
  }
  return false;
}

export function migrateLegacyDataDirectorySync(sourceDirectory: string, targetDirectory: string): void {
  if (containsSymbolicLinkSync(sourceDirectory)) {
    throw new Error("旧数据目录包含符号链接或联接");
  }
  if (hasMeaningfulDataSync(targetDirectory)) {
    throw new Error("用户数据目录已包含文件，拒绝覆盖");
  }

  const parentDirectory = path.dirname(targetDirectory);
  const stagingDirectory = path.join(
    parentDirectory,
    `.${path.basename(targetDirectory)}.legacy-migration-${process.pid}-${Date.now()}`,
  );
  fs.mkdirSync(parentDirectory, { recursive: true });
  fs.rmSync(stagingDirectory, { recursive: true, force: true });

  try {
    fs.cpSync(sourceDirectory, stagingDirectory, {
      recursive: true,
      errorOnExist: true,
      force: false,
      dereference: false,
    });
    if (fs.existsSync(targetDirectory)) {
      fs.rmSync(targetDirectory, { recursive: true, force: true });
    }
    fs.renameSync(stagingDirectory, targetDirectory);
  } catch (error) {
    fs.rmSync(stagingDirectory, { recursive: true, force: true });
    throw error;
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
