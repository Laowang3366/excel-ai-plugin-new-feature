import { randomUUID } from "node:crypto";
import { copyFile, mkdir, readFile, readdir, rename, stat, unlink, writeFile } from "node:fs/promises";
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
  await writeFile(path.join(input.backupRoot, `${id}.json`), `${JSON.stringify(record, null, 2)}\n`, "utf8");
  await pruneOfficeBackups(input.backupRoot);
  return record;
}

export async function listOfficeBackups(backupRoot: string, sourcePath?: string): Promise<OfficeBackupRecord[]> {
  let names: string[];
  try {
    names = await readdir(backupRoot);
  } catch {
    return [];
  }
  const normalizedSource = sourcePath ? path.resolve(sourcePath) : undefined;
  const records = await Promise.all(names.filter((name) => name.endsWith(".json")).map(async (name) => {
    try {
      return JSON.parse(await readFile(path.join(backupRoot, name), "utf8")) as OfficeBackupRecord;
    } catch {
      return undefined;
    }
  }));
  return records
    .filter((record): record is OfficeBackupRecord => Boolean(record?.backupPath && record?.sourcePath))
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
    record = JSON.parse(await readFile(path.join(backupRoot, `${id}.json`), "utf8")) as OfficeBackupRecord;
  } catch {
    throw new Error("Office 事务备份元数据不存在或已损坏");
  }
  if (record.id !== id || !samePath(record.backupPath, backupPath) || !samePath(record.sourcePath, destination)) {
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
  limits: { maxEntries?: number; maxPerSource?: number } = {},
): Promise<void> {
  const maxEntries = limits.maxEntries ?? 500;
  const maxPerSource = limits.maxPerSource ?? 50;
  const records = await listOfficeBackups(backupRoot);
  const keptBySource = new Map<string, number>();
  const expired: OfficeBackupRecord[] = [];
  let kept = 0;
  for (const record of records) {
    const source = path.resolve(record.sourcePath).toLowerCase();
    const sourceCount = keptBySource.get(source) || 0;
    if (kept < maxEntries && sourceCount < maxPerSource) {
      kept++;
      keptBySource.set(source, sourceCount + 1);
    } else {
      expired.push(record);
    }
  }
  await Promise.all(expired.flatMap((record) => [
    unlink(record.backupPath).catch(() => undefined),
    unlink(path.join(backupRoot, `${record.id}.json`)).catch(() => undefined),
  ]));
}
