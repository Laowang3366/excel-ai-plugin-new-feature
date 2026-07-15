import { createHash, randomUUID } from "node:crypto";
import { constants, createReadStream } from "node:fs";
import { promises as fs } from "node:fs";
import path from "node:path";

import Database from "better-sqlite3";

const BACKUP_METADATA_VERSION = 1;
const BACKUP_METADATA_PATTERN = /^analytics-\d{8}T\d{9}Z-[0-9a-f]{8}\.sqlite\.json$/u;

async function sha256File(filePath) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(filePath)) hash.update(chunk);
  return hash.digest("hex");
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function assertHealthyOpenDatabase(db) {
  const quickCheck = db.pragma("quick_check", { simple: true });
  if (quickCheck !== "ok") throw new Error(`SQLite quick_check 失败: ${quickCheck}`);
  const downloadsTable = db.prepare("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'downloads'").get();
  if (!downloadsTable) throw new Error("备份缺少 downloads 表");
}

function assertHealthyAnalyticsDatabase(databasePath) {
  const db = new Database(databasePath, { readonly: true, fileMustExist: true });
  try {
    assertHealthyOpenDatabase(db);
  } finally {
    db.close();
  }
}

function finalizeBackupDatabase(databasePath) {
  const db = new Database(databasePath, { fileMustExist: true });
  try {
    db.pragma("journal_mode = DELETE");
    assertHealthyOpenDatabase(db);
  } finally {
    db.close();
  }
}

async function writeJsonAtomically(filePath, value) {
  const temporaryPath = `${filePath}.tmp-${randomUUID()}`;
  try {
    await fs.writeFile(temporaryPath, `${JSON.stringify(value, null, 2)}\n`, { encoding: "utf8", flag: "wx" });
    await fs.rename(temporaryPath, filePath);
  } catch (error) {
    await fs.rm(temporaryPath, { force: true });
    throw error;
  }
}

async function pruneBackups(outputDir, retain) {
  const entries = await fs.readdir(outputDir, { withFileTypes: true });
  const candidates = [];
  for (const entry of entries) {
    if (!entry.isFile() || !BACKUP_METADATA_PATTERN.test(entry.name)) continue;
    const metadataPath = path.join(outputDir, entry.name);
    try {
      const metadata = JSON.parse(await fs.readFile(metadataPath, "utf8"));
      const backupPath = path.join(outputDir, metadata.backupFile || "");
      if (path.dirname(backupPath) !== outputDir || !(await fileExists(backupPath))) continue;
      candidates.push({ backupPath, metadataPath, createdAt: Date.parse(metadata.createdAt) || 0 });
    } catch {
      // Unknown or incomplete files are preserved for manual inspection.
    }
  }
  candidates.sort((left, right) => right.createdAt - left.createdAt || right.metadataPath.localeCompare(left.metadataPath));
  for (const candidate of candidates.slice(retain)) {
    await fs.rm(candidate.backupPath, { force: true });
    await fs.rm(candidate.metadataPath, { force: true });
  }
}

export async function verifyAnalyticsBackup(backupPath) {
  const resolvedBackupPath = path.resolve(backupPath);
  const metadataPath = `${resolvedBackupPath}.json`;
  const metadata = JSON.parse(await fs.readFile(metadataPath, "utf8"));
  if (metadata.schemaVersion !== BACKUP_METADATA_VERSION) throw new Error("备份元数据版本不受支持");
  if (metadata.backupFile !== path.basename(resolvedBackupPath)) throw new Error("备份元数据文件名不匹配");
  const stat = await fs.stat(resolvedBackupPath);
  if (metadata.size !== stat.size) throw new Error("备份大小校验失败");
  if (metadata.sha256 !== await sha256File(resolvedBackupPath)) throw new Error("备份 SHA-256 校验失败");
  assertHealthyAnalyticsDatabase(resolvedBackupPath);
  return { backupPath: resolvedBackupPath, metadataPath, metadata };
}

export async function createAnalyticsBackup({ sourcePath, outputDir, retain = 14, now = Date.now }) {
  if (!Number.isInteger(retain) || retain < 1 || retain > 365) throw new Error("retain 必须是 1-365 之间的整数");
  const resolvedSourcePath = path.resolve(sourcePath);
  const resolvedOutputDir = path.resolve(outputDir);
  await fs.mkdir(resolvedOutputDir, { recursive: true });
  const createdAt = new Date(now()).toISOString();
  const timestamp = createdAt.replace(/[-:.]/gu, "");
  const backupFile = `analytics-${timestamp}-${randomUUID().slice(0, 8)}.sqlite`;
  const backupPath = path.join(resolvedOutputDir, backupFile);
  const temporaryBackupPath = `${backupPath}.tmp`;
  const metadataPath = `${backupPath}.json`;
  try {
    const source = new Database(resolvedSourcePath, { readonly: true, fileMustExist: true });
    try {
      await source.backup(temporaryBackupPath);
    } finally {
      source.close();
    }
    finalizeBackupDatabase(temporaryBackupPath);
    const stat = await fs.stat(temporaryBackupPath);
    const metadata = {
      schemaVersion: BACKUP_METADATA_VERSION,
      createdAt,
      sourceFile: path.basename(resolvedSourcePath),
      backupFile,
      size: stat.size,
      sha256: await sha256File(temporaryBackupPath),
    };
    await fs.rename(temporaryBackupPath, backupPath);
    await writeJsonAtomically(metadataPath, metadata);
    await verifyAnalyticsBackup(backupPath);
    await pruneBackups(resolvedOutputDir, retain);
    return { backupPath, metadataPath, metadata };
  } catch (error) {
    await fs.rm(temporaryBackupPath, { force: true });
    await fs.rm(backupPath, { force: true });
    await fs.rm(metadataPath, { force: true });
    throw error;
  }
}

export async function restoreAnalyticsBackup({ backupPath, targetPath }) {
  const verified = await verifyAnalyticsBackup(backupPath);
  const resolvedTargetPath = path.resolve(targetPath);
  if (await fileExists(resolvedTargetPath)) throw new Error("恢复目标已存在；请先停止服务并移走当前数据库");
  await fs.mkdir(path.dirname(resolvedTargetPath), { recursive: true });
  const temporaryTargetPath = `${resolvedTargetPath}.restore-${randomUUID()}`;
  try {
    await fs.copyFile(verified.backupPath, temporaryTargetPath, constants.COPYFILE_EXCL);
    assertHealthyAnalyticsDatabase(temporaryTargetPath);
    if (await sha256File(temporaryTargetPath) !== verified.metadata.sha256) {
      throw new Error("恢复副本 SHA-256 校验失败");
    }
    const handle = await fs.open(temporaryTargetPath, "r+");
    try {
      await handle.sync();
    } finally {
      await handle.close();
    }
    await fs.rename(temporaryTargetPath, resolvedTargetPath);
    return { targetPath: resolvedTargetPath, metadata: verified.metadata };
  } catch (error) {
    await fs.rm(temporaryTargetPath, { force: true });
    throw error;
  }
}
