import { createHash } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";

import { SqliteStore } from "../agent/knowledge";
import { StateRuntimeStore } from "../agent/memory/stateRuntimeStore";
import { isPathInside, normalizePathForCompare, pathExists } from "./settingsDataPath";

export interface PreparedDataPathMigration {
  currentDataPath: string;
  targetDataPath: string;
  stageDataPath: string;
  targetExisted: boolean;
}

export async function prepareDataPathMigration(
  currentDataPath: string,
  targetDataPath: string,
): Promise<PreparedDataPathMigration> {
  const current = path.resolve(currentDataPath);
  const target = path.resolve(targetDataPath);
  assertMigrationPaths(current, target);

  const targetExisted = await pathExists(target);
  if (targetExisted) {
    const stat = await fs.promises.stat(target);
    if (!stat.isDirectory()) throw new Error("新数据路径必须是目录");
    const entries = await fs.promises.readdir(target);
    if (entries.length > 0) {
      throw new Error("为避免覆盖现有文件，新数据目录必须为空");
    }
  }

  const parent = path.dirname(target);
  await fs.promises.mkdir(parent, { recursive: true });
  const stageDataPath = path.join(
    parent,
    `.${path.basename(target)}.wengge-migration-${process.pid}-${Date.now()}`,
  );
  await fs.promises.rm(stageDataPath, { recursive: true, force: true });
  await fs.promises.mkdir(stageDataPath, { recursive: true });

  try {
    await copyDirectoryVerified(current, stageDataPath, "");
    return {
      currentDataPath: current,
      targetDataPath: target,
      stageDataPath,
      targetExisted,
    };
  } catch (error) {
    await fs.promises.rm(stageDataPath, { recursive: true, force: true }).catch(() => {});
    throw error;
  }
}

export async function commitPreparedDataPathMigration(
  prepared: PreparedDataPathMigration,
): Promise<void> {
  if (await pathExists(prepared.targetDataPath)) {
    const entries = await fs.promises.readdir(prepared.targetDataPath);
    if (entries.length > 0) {
      throw new Error("提交迁移前目标目录不再为空，已取消切换");
    }
    await fs.promises.rmdir(prepared.targetDataPath);
  }
  await fs.promises.rename(prepared.stageDataPath, prepared.targetDataPath);
}

export async function cleanupPreparedDataPathMigration(
  prepared: PreparedDataPathMigration | null,
  removeCommittedTarget: boolean,
): Promise<void> {
  if (!prepared) return;
  await fs.promises.rm(prepared.stageDataPath, { recursive: true, force: true }).catch(() => {});
  if (removeCommittedTarget) {
    await fs.promises.rm(prepared.targetDataPath, { recursive: true, force: true }).catch(() => {});
    if (prepared.targetExisted) {
      await fs.promises.mkdir(prepared.targetDataPath, { recursive: true }).catch(() => {});
    }
  }
}

export async function validateStagedDataPath(stageDataPath: string): Promise<void> {
  const stateStore = new StateRuntimeStore(
    path.join(stageDataPath, "sessions", "state-runtime"),
  );
  await stateStore.init();
  await stateStore.close();

  const knowledgeStore = new SqliteStore(
    path.join(stageDataPath, "knowledge", "knowledge.db"),
  );
  try {
    await knowledgeStore.init();
  } finally {
    knowledgeStore.close();
  }
}

function assertMigrationPaths(current: string, target: string): void {
  if (normalizePathForCompare(current) === normalizePathForCompare(target)) return;
  if (target.startsWith("\\\\")) {
    throw new Error("默认不允许把数据目录迁移到 UNC 或网络共享路径");
  }
  if (path.dirname(target) === target) {
    throw new Error("请选择磁盘根目录下的专用子目录作为数据目录");
  }
  if (isPathInside(current, target)) {
    throw new Error("新数据目录不能位于当前数据目录内部");
  }
  if (isPathInside(target, current)) {
    throw new Error("新数据目录不能包含当前数据目录");
  }
}

async function copyDirectoryVerified(
  sourceDir: string,
  targetDir: string,
  relativeDir: string,
): Promise<void> {
  const entries = await fs.promises.readdir(sourceDir, { withFileTypes: true });
  for (const entry of entries) {
    const relativePath = path.join(relativeDir, entry.name);
    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = path.join(targetDir, entry.name);
    if (entry.isSymbolicLink()) {
      throw new Error(`数据目录包含不允许迁移的符号链接或联接: ${relativePath}`);
    }
    if (entry.isDirectory()) {
      await fs.promises.mkdir(targetPath, { recursive: true });
      await copyDirectoryVerified(sourcePath, targetPath, relativePath);
      continue;
    }
    if (!entry.isFile()) {
      throw new Error(`数据目录包含不支持的文件类型: ${relativePath}`);
    }

    await fs.promises.copyFile(sourcePath, targetPath, fs.constants.COPYFILE_EXCL);
    if (isVolatilePath(relativePath)) continue;
    const [sourceHash, targetHash] = await Promise.all([
      hashFile(sourcePath),
      hashFile(targetPath),
    ]);
    if (sourceHash !== targetHash) {
      throw new Error(`迁移校验失败，文件内容不一致: ${relativePath}`);
    }
  }
}

function isVolatilePath(relativePath: string): boolean {
  const root = relativePath.split(/[\\/]/, 1)[0]?.toLowerCase();
  return root === "logs" || root === "temp";
}

function hashFile(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}
