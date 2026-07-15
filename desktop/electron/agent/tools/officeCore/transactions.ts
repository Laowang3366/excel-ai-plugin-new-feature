import { randomUUID } from "node:crypto";
import {
  copyFile,
  mkdir,
  readFile,
  readdir,
  rename,
  stat,
  unlink,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

export interface OfficeBackupRecord {
  id: string;
  app: "excel" | "word" | "presentation";
  operation: string;
  sourcePath: string;
  backupPath: string;
  createdAt: string;
  size: number;
}

export interface OfficeBackupPruneResult {
  deletedRecords: number;
  deletedFiles: number;
  reclaimedBytes: number;
  errors: string[];
}

export async function createOfficeBackup(input: {
  backupRoot: string;
  app: OfficeBackupRecord["app"];
  operation: string;
  sourcePath: string;
}): Promise<OfficeBackupRecord> {
  const source = path.resolve(input.sourcePath);
  const sourceStat = await stat(source);
  if (!sourceStat.isFile()) throw new Error(`Office 事务源文件不存在: ${source}`);

  await mkdir(input.backupRoot, { recursive: true });
  const id = `${Date.now()}-${randomUUID()}`;
  const backupPath = path.join(input.backupRoot, `${id}${path.extname(source)}`);
  const record: OfficeBackupRecord = {
    id,
    app: input.app,
    operation: input.operation,
    sourcePath: source,
    backupPath,
    createdAt: new Date().toISOString(),
    size: sourceStat.size,
  };
  await copyFile(source, backupPath);
  await writeFile(
    path.join(input.backupRoot, `${id}.json`),
    `${JSON.stringify(record, null, 2)}\n`,
    "utf8",
  );
  await pruneOfficeBackups(input.backupRoot, { protectedIds: [id] });
  return record;
}

export async function listOfficeBackups(
  backupRoot: string,
  sourcePath?: string,
): Promise<OfficeBackupRecord[]> {
  let names: string[];
  try {
    names = await readdir(backupRoot);
  } catch {
    return [];
  }
  const normalizedSource = sourcePath ? path.resolve(sourcePath) : undefined;
  const root = path.resolve(backupRoot);
  const records = await Promise.all(
    names
      .filter((name) => name.endsWith(".json"))
      .map(async (name) => {
        try {
          const record = JSON.parse(
            await readFile(path.join(root, name), "utf8"),
          ) as OfficeBackupRecord;
          return isManagedBackupRecord(root, name, record) ? record : undefined;
        } catch {
          return undefined;
        }
      }),
  );
  return records
    .filter((record): record is OfficeBackupRecord =>
      Boolean(record?.backupPath && record?.sourcePath),
    )
    .filter((record) => !normalizedSource || path.resolve(record.sourcePath) === normalizedSource)
    .sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function restoreOfficeBackup(input: {
  backupRoot: string;
  backupPath: string;
  destinationPath: string;
}): Promise<void> {
  const backupRoot = path.resolve(input.backupRoot);
  const backupPath = path.resolve(input.backupPath);
  const relative = path.relative(backupRoot, backupPath);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error("备份文件不在受控 Office 事务目录中");
  }
  const destination = path.resolve(input.destinationPath);
  const id = path.parse(backupPath).name;
  let record: OfficeBackupRecord;
  try {
    record = JSON.parse(
      await readFile(path.join(backupRoot, `${id}.json`), "utf8"),
    ) as OfficeBackupRecord;
  } catch {
    throw new Error("Office 事务备份元数据不存在或已损坏");
  }
  if (
    record.id !== id ||
    !samePath(record.backupPath, backupPath) ||
    !samePath(record.sourcePath, destination)
  ) {
    throw new Error("Office 事务备份与目标文件不匹配");
  }
  const temporary = `${destination}.${randomUUID()}.restore.tmp`;
  await copyFile(backupPath, temporary);
  await rename(temporary, destination);
}

function samePath(left: string, right: string): boolean {
  const normalizedLeft = path.resolve(left);
  const normalizedRight = path.resolve(right);
  return process.platform === "win32"
    ? normalizedLeft.toLowerCase() === normalizedRight.toLowerCase()
    : normalizedLeft === normalizedRight;
}

export async function pruneOfficeBackups(
  backupRoot: string,
  limits: {
    maxEntries?: number;
    maxPerSource?: number;
    maxAgeDays?: number;
    maxBytes?: number;
    now?: number;
    protectedIds?: string[];
  } = {},
): Promise<OfficeBackupPruneResult> {
  const maxEntries = limits.maxEntries ?? 500;
  const maxPerSource = limits.maxPerSource ?? 50;
  const maxAgeDays = limits.maxAgeDays ?? 30;
  const maxBytes = limits.maxBytes ?? 2 * 1024 * 1024 * 1024;
  const cutoff = (limits.now ?? Date.now()) - maxAgeDays * 24 * 60 * 60 * 1000;
  const protectedIds = new Set(limits.protectedIds || []);
  const records = await listOfficeBackups(backupRoot);
  const keptBySource = new Map<string, number>();
  const expired: OfficeBackupRecord[] = [];
  let kept = 0;
  let keptBytes = 0;
  for (const record of records) {
    const source = path.resolve(record.sourcePath).toLowerCase();
    const sourceCount = keptBySource.get(source) || 0;
    const createdAt = Date.parse(record.createdAt);
    const protectedRecord = protectedIds.has(record.id);
    if (
      protectedRecord ||
      (createdAt >= cutoff &&
        kept < maxEntries &&
        sourceCount < maxPerSource &&
        keptBytes + record.size <= maxBytes)
    ) {
      kept++;
      keptBytes += record.size;
      keptBySource.set(source, sourceCount + 1);
    } else {
      expired.push(record);
    }
  }

  const result: OfficeBackupPruneResult = {
    deletedRecords: 0,
    deletedFiles: 0,
    reclaimedBytes: 0,
    errors: [],
  };
  for (const record of expired) {
    let backupDeleted = false;
    try {
      await unlink(record.backupPath);
      backupDeleted = true;
      result.deletedFiles++;
      result.reclaimedBytes += record.size;
    } catch (error) {
      if (!isFileNotFound(error)) result.errors.push(errorMessage(error));
    }
    try {
      await unlink(path.join(path.resolve(backupRoot), `${record.id}.json`));
      result.deletedFiles++;
    } catch (error) {
      if (!isFileNotFound(error)) result.errors.push(errorMessage(error));
    }
    if (backupDeleted) result.deletedRecords++;
  }
  return result;
}

function isManagedBackupRecord(
  backupRoot: string,
  metadataName: string,
  record: OfficeBackupRecord,
): boolean {
  const id = path.basename(metadataName, ".json");
  if (
    record?.id !== id ||
    typeof record.backupPath !== "string" ||
    typeof record.sourcePath !== "string" ||
    !Number.isFinite(record.size) ||
    record.size < 0 ||
    !Number.isFinite(Date.parse(record.createdAt))
  )
    return false;

  const backupPath = path.resolve(record.backupPath);
  if (!samePath(path.dirname(backupPath), backupRoot)) return false;
  const backupName = path.basename(backupPath);
  return backupName.startsWith(`${id}.`) && backupName !== `${id}.json`;
}

function isFileNotFound(error: unknown): boolean {
  return Boolean(error && typeof error === "object" && "code" in error && error.code === "ENOENT");
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
