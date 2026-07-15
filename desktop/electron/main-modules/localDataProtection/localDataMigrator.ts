import * as fs from "node:fs";
import * as path from "node:path";
import { createLogger } from "../../shared/logger";
import { pathExists } from "../settingsDataPath";
import {
  clearMigrationJournal,
  fixedTransactionPaths,
  writeMigrationJournal,
} from "./migrationJournal";
import {
  transformArchiveFiles,
  transformJsonlFiles,
  transformStagePayloads,
} from "./localDataMigratorOps";
import { validateAllEncrypted } from "./localDataMigratorValidate";
import type { PayloadProtection } from "./payloadProtection";

const logger = createLogger("LocalDataMigrator");
const MARKER_REL = path.join("sessions", ".local-data-protection.json");

export interface LocalDataProtectionMarker {
  formatVersion: 1;
  contentKeyId: number;
  migratedAt: string;
}

export interface FinalizeResult {
  cleared: boolean;
  backupRoot: string;
  error?: string;
}

export interface AtomicMigrationHandle {
  dataRoot: string;
  backupRoot: string;
  stageRoot: string;
  targetKeyId: number;
  previousKeyId: number;
  finalize(): Promise<FinalizeResult>;
  rollback(): Promise<void>;
  markCommitted(): void;
}

function protectionMarkerPath(dataRoot: string): string {
  return path.join(dataRoot, MARKER_REL);
}

export function readProtectionMarker(dataRoot: string): LocalDataProtectionMarker | null {
  const markerPath = protectionMarkerPath(dataRoot);
  if (!fs.existsSync(markerPath)) return null;
  return JSON.parse(fs.readFileSync(markerPath, "utf8")) as LocalDataProtectionMarker;
}

function writeMarker(dataRoot: string, contentKeyId: number): void {
  const markerPath = protectionMarkerPath(dataRoot);
  fs.mkdirSync(path.dirname(markerPath), { recursive: true });
  const marker: LocalDataProtectionMarker = {
    formatVersion: 1,
    contentKeyId,
    migratedAt: new Date().toISOString(),
  };
  const tempPath = `${markerPath}.${process.pid}.tmp`;
  fs.writeFileSync(tempPath, `${JSON.stringify(marker, null, 2)}\n`, "utf8");
  fs.renameSync(tempPath, markerPath);
}

function isVolatilePath(relativePath: string): boolean {
  const root = relativePath.split(/[\\/]/, 1)[0]?.toLowerCase();
  return root === "logs" || root === "temp";
}

async function hashFile(filePath: string): Promise<string> {
  const { createHash } = await import("node:crypto");
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
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
    if (!entry.isFile()) throw new Error(`数据目录包含不支持的文件类型: ${relativePath}`);
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

/**
 * Stage → transform → full validate → swap, keep backup until finalize().
 * Uses fixed sibling paths + migration journal for crash recovery.
 */
export async function migrateLocalDataAtomically(options: {
  dataRoot: string;
  protection: PayloadProtection;
  targetKeyId: number;
  previousKeyId: number;
  kind: "encrypt" | "rotate";
  userDataPath?: string;
  force?: boolean;
}): Promise<AtomicMigrationHandle | { migrated: false; contentKeyId: number }> {
  const dataRoot = path.resolve(options.dataRoot);
  const marker = fs.existsSync(dataRoot) ? readProtectionMarker(dataRoot) : null;
  if (!options.force && marker?.contentKeyId === options.targetKeyId) {
    return { migrated: false, contentKeyId: options.targetKeyId };
  }
  if (!fs.existsSync(dataRoot)) {
    writeMarker(dataRoot, options.targetKeyId);
    return {
      dataRoot,
      backupRoot: "",
      stageRoot: "",
      targetKeyId: options.targetKeyId,
      previousKeyId: options.previousKeyId,
      finalize: async () => ({ cleared: true, backupRoot: "" }),
      rollback: async () => {},
      markCommitted: () => {},
    };
  }

  const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
  if (await pathExists(backupRoot)) {
    throw new Error("检测到未完成的加密备份目录，请先完成恢复或清理后再迁移（.wengge-ldp-backup）");
  }
  const journalBase = {
    formatVersion: 1 as const,
    kind: options.kind,
    dataRoot,
    stageRoot,
    backupRoot,
    targetKeyId: options.targetKeyId,
    previousKeyId: options.previousKeyId,
    updatedAt: new Date().toISOString(),
  };

  try {
    if (await pathExists(stageRoot)) {
      await fs.promises.rm(stageRoot, { recursive: true, force: false });
    }
  } catch (error) {
    throw new Error(
      `无法清理加密 staging 目录: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  await fs.promises.mkdir(stageRoot, { recursive: true });
  writeMigrationJournal({ ...journalBase, phase: "staging" }, options.userDataPath);

  try {
    await copyDirectoryVerified(dataRoot, stageRoot, "");
    await transformJsonlFiles(stageRoot, options.protection, options.targetKeyId);
    await transformArchiveFiles(stageRoot, options.protection, options.targetKeyId);
    transformStagePayloads(stageRoot, options.protection, options.targetKeyId);
    await validateAllEncrypted(stageRoot, options.protection, options.targetKeyId);
    writeMarker(stageRoot, options.targetKeyId);

    writeMigrationJournal({ ...journalBase, phase: "swapping" }, options.userDataPath);
    await fs.promises.rename(dataRoot, backupRoot);
    try {
      await fs.promises.rename(stageRoot, dataRoot);
    } catch (error) {
      await fs.promises.rename(backupRoot, dataRoot);
      clearMigrationJournal(options.userDataPath);
      throw error;
    }
    writeMigrationJournal({ ...journalBase, phase: "swapped" }, options.userDataPath);

    logger.info("Local data encryption swap complete; backup retained until finalize", {
      contentKeyId: options.targetKeyId,
      dataRoot,
      backupRoot,
    });

    return {
      dataRoot,
      backupRoot,
      stageRoot,
      targetKeyId: options.targetKeyId,
      previousKeyId: options.previousKeyId,
      markCommitted: () => {
        writeMigrationJournal({ ...journalBase, phase: "committed" }, options.userDataPath);
      },
      finalize: async () => {
        writeMigrationJournal({ ...journalBase, phase: "finalizing" }, options.userDataPath);
        if (!backupRoot || !(await pathExists(backupRoot))) {
          clearMigrationJournal(options.userDataPath);
          return { cleared: true, backupRoot };
        }
        try {
          await fs.promises.rm(backupRoot, { recursive: true, force: false });
          clearMigrationJournal(options.userDataPath);
          return { cleared: true, backupRoot };
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          return { cleared: false, backupRoot, error: message };
        }
      },
      rollback: async () => {
        if (await pathExists(dataRoot)) {
          await fs.promises.rm(dataRoot, { recursive: true, force: false });
        }
        if (await pathExists(backupRoot)) {
          await fs.promises.rename(backupRoot, dataRoot);
        }
        if (await pathExists(stageRoot)) {
          await fs.promises.rm(stageRoot, { recursive: true, force: false });
        }
        clearMigrationJournal(options.userDataPath);
      },
    };
  } catch (error) {
    if (await pathExists(stageRoot)) {
      await fs.promises.rm(stageRoot, { recursive: true, force: false }).catch(() => {});
    }
    if ((await pathExists(backupRoot)) && !(await pathExists(dataRoot))) {
      await fs.promises.rename(backupRoot, dataRoot);
    }
    clearMigrationJournal(options.userDataPath);
    throw error;
  }
}

export {
  fixedTransactionPaths,
  readMigrationJournal,
  clearMigrationJournal,
  type MigrationJournalRecord,
} from "./migrationJournal";
