import * as fs from "node:fs";
import * as path from "node:path";

export { USER_DATA_ERASE_CONFIRMATION } from "../shared/userDataEraseContract";

const MANAGED_DATA_DIRECTORIES = [
  "sessions",
  "knowledge",
  "office-backups",
  "office-automation",
  "logs",
  "temp",
] as const;

export interface UserDataEraseReport {
  erasedCategories: string[];
  errors: string[];
}

export async function eraseManagedUserData(
  dataPath: string,
  options: {
    removeDirectory?: (directory: string) => Promise<void>;
  } = {},
): Promise<UserDataEraseReport> {
  const root = path.resolve(dataPath);
  await assertSafeDataRoot(root);
  const removeDirectory =
    options.removeDirectory ??
    ((directory) => fs.promises.rm(directory, { recursive: true, force: false }));
  const report: UserDataEraseReport = { erasedCategories: [], errors: [] };

  for (const category of MANAGED_DATA_DIRECTORIES) {
    const target = path.join(root, category);
    try {
      assertDirectChild(root, target);
      const info = await fs.promises.lstat(target).catch((error: NodeJS.ErrnoException) => {
        if (error.code === "ENOENT") return null;
        throw error;
      });
      if (info?.isSymbolicLink()) {
        throw new Error("拒绝删除符号链接或联接");
      }
      if (info && !info.isDirectory()) {
        throw new Error("受管数据路径不是目录");
      }
      if (info) await removeDirectory(target);
      report.erasedCategories.push(category);
    } catch (error) {
      report.errors.push(`${category}: ${errorMessage(error)}`);
    }
  }
  return report;
}

async function assertSafeDataRoot(root: string): Promise<void> {
  if (path.dirname(root) === root) throw new Error("拒绝擦除磁盘根目录");
  const info = await fs.promises.lstat(root);
  if (!info.isDirectory() || info.isSymbolicLink()) {
    throw new Error("数据根目录必须是非链接目录");
  }
}

function assertDirectChild(root: string, target: string): void {
  const relative = path.relative(root, path.resolve(target));
  if (!relative || relative.includes(path.sep) || path.isAbsolute(relative)) {
    throw new Error("拒绝删除受管数据根之外的路径");
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
