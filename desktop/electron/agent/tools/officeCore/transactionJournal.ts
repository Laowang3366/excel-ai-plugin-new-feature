import { createHash, randomUUID } from "node:crypto";
import { copyFile, mkdir, readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";

import type { OfficeActionBridge } from "../contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "./types";

export type OfficeTransactionStatus = "pending" | "applied" | "undone" | "failed" | "conflicted";

export interface OfficeTransactionSnapshot {
  filePath: string;
  existed: boolean;
  snapshotPath?: string;
  beforeHash?: string;
  afterExisted?: boolean;
  afterSnapshotPath?: string;
  afterHash?: string;
}

export interface OfficeTransactionConflict {
  filePath: string;
  expected: "before" | "after";
  reason: string;
}

export interface OfficeTransactionRestoreFile {
  filePath: string;
  existed: boolean;
  snapshotPath?: string;
}

export interface OfficeTransactionRestoreOptions {
  force?: boolean;
  prepareFiles?: (filePaths: string[]) => Promise<unknown>;
  restoreFiles?: (files: OfficeTransactionRestoreFile[]) => Promise<unknown>;
}

export interface OfficeTransactionRecord {
  id: string;
  workflowId?: string;
  status: OfficeTransactionStatus;
  createdAt: string;
  updatedAt: string;
  steps: OfficeActionInput[];
  results: OfficeActionResult[];
  snapshots: OfficeTransactionSnapshot[];
  artifacts: string[];
  changes: OfficeActionResult["changes"];
  conflicts?: OfficeTransactionConflict[];
  conflictBaseStatus?: Exclude<OfficeTransactionStatus, "conflicted">;
  error?: string;
}

export async function beginOfficeTransaction(input: {
  root: string;
  steps: OfficeActionInput[];
  workflowId?: string;
}): Promise<OfficeTransactionRecord> {
  const id = randomUUID();
  const transactionDir = transactionDirectory(input.root, id);
  const snapshotDir = path.join(transactionDir, "snapshots");
  await mkdir(snapshotDir, { recursive: true });
  const snapshots: OfficeTransactionSnapshot[] = [];
  for (const [index, filePath] of listOfficeTransactionPaths(input.steps).entries()) {
    const snapshot = await snapshotPath(filePath, snapshotDir, index);
    if (snapshot) snapshots.push(snapshot);
  }
  const now = new Date().toISOString();
  const record: OfficeTransactionRecord = {
    id,
    workflowId: input.workflowId,
    status: "pending",
    createdAt: now,
    updatedAt: now,
    steps: input.steps,
    results: [],
    snapshots,
    artifacts: [],
    changes: [],
  };
  await saveOfficeTransaction(input.root, record);
  return record;
}

export async function saveOfficeTransaction(root: string, record: OfficeTransactionRecord): Promise<void> {
  record.updatedAt = new Date().toISOString();
  const directory = transactionDirectory(root, record.id);
  await mkdir(directory, { recursive: true });
  const destination = path.join(directory, "transaction.json");
  const temporary = path.join(directory, `.transaction.${randomUUID()}.tmp`);
  await writeFile(temporary, `${JSON.stringify(record, null, 2)}\n`, "utf8");
  await rm(destination, { force: true });
  await rename(temporary, destination);
}

export async function getOfficeTransaction(root: string, id: string): Promise<OfficeTransactionRecord> {
  validateRecordId(id);
  const record = JSON.parse(await readFile(path.join(transactionDirectory(root, id), "transaction.json"), "utf8")) as OfficeTransactionRecord;
  if (record.id !== id || !Array.isArray(record.steps) || !Array.isArray(record.snapshots)) {
    throw new Error("Office 事务记录已损坏");
  }
  return record;
}

export async function listOfficeTransactions(root: string): Promise<OfficeTransactionRecord[]> {
  let names: string[];
  try { names = await readdir(root); } catch { return []; }
  const records = await Promise.all(names.map(async (name) => {
    try { return await getOfficeTransaction(root, name); } catch { return undefined; }
  }));
  return records
    .filter((record): record is OfficeTransactionRecord => Boolean(record))
    .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
}

export async function recordOfficeTransactionResult(
  root: string,
  record: OfficeTransactionRecord,
  result: OfficeActionResult,
): Promise<void> {
  record.results.push(result);
  record.changes.push(...result.changes);
  for (const artifact of collectResultArtifacts(result)) {
    if (!record.artifacts.includes(artifact)) record.artifacts.push(artifact);
  }
  await saveOfficeTransaction(root, record);
}

export async function finalizeOfficeTransaction(
  root: string,
  record: OfficeTransactionRecord,
): Promise<OfficeTransactionRecord> {
  const untracked = record.artifacts.filter((artifact) =>
    !record.snapshots.some((snapshot) => samePath(snapshot.filePath, artifact))
  );
  if (untracked.length > 0) {
    record.status = "failed";
    record.error = `Office 事务存在未声明产物，无法保证完整恢复: ${untracked.join(", ")}`;
    await saveOfficeTransaction(root, record);
    throw new Error(record.error);
  }

  const afterDir = path.join(transactionDirectory(root, record.id), "after");
  await mkdir(afterDir, { recursive: true });
  for (const [index, snapshot] of record.snapshots.entries()) {
    const after = await captureFileState(snapshot.filePath, afterDir, index);
    snapshot.afterExisted = after.existed;
    snapshot.afterSnapshotPath = after.snapshotPath;
    snapshot.afterHash = after.hash;
  }
  record.status = "applied";
  record.conflicts = [];
  record.conflictBaseStatus = undefined;
  record.error = undefined;
  await saveOfficeTransaction(root, record);
  return record;
}

export async function undoOfficeTransaction(
  root: string,
  id: string,
  options: OfficeTransactionRestoreOptions = {},
): Promise<OfficeTransactionRecord> {
  const record = await getOfficeTransaction(root, id);
  if (record.status === "undone") return record;
  if (options.prepareFiles) await options.prepareFiles(record.snapshots.map((snapshot) => snapshot.filePath));
  if (hasAfterState(record) && !options.force) {
    const conflicts = await detectTransactionConflicts(record, "after");
    if (conflicts.length > 0) return saveConflict(root, record, conflicts, "撤销");
  }
  await restoreSnapshots(root, record.id, [...record.snapshots].reverse(), "before", options.restoreFiles);
  record.status = "undone";
  record.conflicts = [];
  record.conflictBaseStatus = undefined;
  record.error = undefined;
  await saveOfficeTransaction(root, record);
  return record;
}

export async function redoOfficeTransaction(
  root: string,
  id: string,
  bridge?: OfficeActionBridge,
  options: OfficeTransactionRestoreOptions = {},
): Promise<OfficeTransactionRecord> {
  const record = await getOfficeTransaction(root, id);
  const canRedo = record.status === "undone"
    || (record.status === "conflicted" && record.conflictBaseStatus === "undone");
  if (!canRedo) throw new Error("只有已撤销的 Office 事务可以重新执行");
  if (hasAfterState(record)) {
    if (options.prepareFiles) await options.prepareFiles(record.snapshots.map((snapshot) => snapshot.filePath));
    if (!options.force) {
      const conflicts = await detectTransactionConflicts(record, "before");
      if (conflicts.length > 0) return saveConflict(root, record, conflicts, "重做");
    }
    await restoreSnapshots(root, record.id, record.snapshots, "after", options.restoreFiles);
    record.status = "applied";
    record.conflicts = [];
    record.conflictBaseStatus = undefined;
    record.error = undefined;
    await saveOfficeTransaction(root, record);
    return record;
  }
  if (!bridge) throw new Error("旧版 Office 事务缺少 after 快照，无法确定性重做");
  record.status = "pending";
  record.results = [];
  record.artifacts = [];
  record.changes = [];
  record.error = undefined;
  await saveOfficeTransaction(root, record);
  for (const step of record.steps) {
    const result = await bridge.executeAction(step);
    await recordOfficeTransactionResult(root, record, result);
    if (result.status === "done") continue;
    record.status = "failed";
    record.error = result.error || result.summary;
    await saveOfficeTransaction(root, record);
    await undoOfficeTransaction(root, record.id, { force: true });
    throw new Error(record.error);
  }
  record.status = "applied";
  await saveOfficeTransaction(root, record);
  return record;
}

export function listOfficeTransactionPaths(steps: OfficeActionInput[]): string[] {
  const paths = new Map<string, string>();
  for (const step of steps) {
    addPath(paths, step.filePath);
    addPath(paths, step.outputPath);
    const params = step.params || {};
    for (const key of ["outputPath", "wordOutputPath", "presentationOutputPath"] as const) {
      if (typeof params[key] === "string") addPath(paths, params[key]);
    }
    if (step.operation === "buildReportPackage") {
      const outputDirectory = typeof params.outputDirectory === "string" ? params.outputDirectory : step.outputPath;
      if (outputDirectory) {
        const baseName = typeof params.baseName === "string"
          ? params.baseName
          : `${path.basename(step.filePath || "report", path.extname(step.filePath || ""))}-报告`;
        addPath(paths, path.join(outputDirectory, `${baseName}.docx`));
        addPath(paths, path.join(outputDirectory, `${baseName}.pptx`));
      }
    }
  }
  return [...paths.values()];
}

function collectResultArtifacts(result: OfficeActionResult): string[] {
  const artifacts = new Map<string, string>();
  if (result.outputPath && result.filePath && !samePath(result.outputPath, result.filePath)) addPath(artifacts, result.outputPath);
  for (const change of result.changes) {
    if (change.target && path.isAbsolute(change.target) && path.extname(change.target)) addPath(artifacts, change.target);
  }
  return [...artifacts.values()];
}

async function snapshotPath(filePath: string, snapshotDir: string, index: number): Promise<OfficeTransactionSnapshot | undefined> {
  const state = await captureFileState(filePath, snapshotDir, index);
  if (state.directory) return undefined;
  return {
    filePath: path.resolve(filePath),
    existed: state.existed,
    snapshotPath: state.snapshotPath,
    beforeHash: state.hash,
  };
}

async function restoreSnapshots(
  root: string,
  id: string,
  snapshots: OfficeTransactionSnapshot[],
  phase: "before" | "after",
  restoreFiles?: (files: OfficeTransactionRestoreFile[]) => Promise<unknown>,
): Promise<void> {
  if (!restoreFiles) {
    await restoreSnapshotsAtomically(root, id, snapshots, phase);
    return;
  }
  const files = snapshots.map((snapshot) => restoreFileDescriptor(root, id, snapshot, phase));
  await restoreFiles(files);
}

interface StagedRestore {
  destination: string;
  stagedPath?: string;
  rollbackPath?: string;
  committed: boolean;
}

async function restoreSnapshotsAtomically(
  root: string,
  id: string,
  snapshots: OfficeTransactionSnapshot[],
  phase: "before" | "after",
): Promise<void> {
  const staged: StagedRestore[] = [];
  try {
    for (const snapshot of snapshots) {
      const descriptor = restoreFileDescriptor(root, id, snapshot, phase);
      const destination = path.resolve(descriptor.filePath);
      await mkdir(path.dirname(destination), { recursive: true });
      let stagedPath: string | undefined;
      if (descriptor.existed) {
        stagedPath = `${destination}.${randomUUID()}.transaction.stage`;
        await copyFile(descriptor.snapshotPath!, stagedPath);
      }
      staged.push({ destination, stagedPath, committed: false });
    }

    for (const entry of staged) {
      if (await fileExists(entry.destination)) {
        entry.rollbackPath = `${entry.destination}.${randomUUID()}.transaction.rollback`;
        await rename(entry.destination, entry.rollbackPath);
      }
      entry.committed = true;
      if (entry.stagedPath) {
        await rename(entry.stagedPath, entry.destination);
        entry.stagedPath = undefined;
      }
    }
  } catch (error) {
    for (const entry of [...staged].reverse()) {
      if (!entry.committed) continue;
      await rm(entry.destination, { force: true }).catch(() => undefined);
      if (entry.rollbackPath) await rename(entry.rollbackPath, entry.destination).catch(() => undefined);
      entry.rollbackPath = undefined;
    }
    throw error;
  } finally {
    for (const entry of staged) {
      if (entry.stagedPath) await rm(entry.stagedPath, { force: true }).catch(() => undefined);
      if (entry.rollbackPath) await rm(entry.rollbackPath, { force: true }).catch(() => undefined);
    }
  }
}

function restoreFileDescriptor(
  root: string,
  id: string,
  snapshot: OfficeTransactionSnapshot,
  phase: "before" | "after",
): OfficeTransactionRestoreFile {
  const existed = phase === "before" ? snapshot.existed : snapshot.afterExisted === true;
  const snapshotPath = phase === "before" ? snapshot.snapshotPath : snapshot.afterSnapshotPath;
  if (!existed) return { filePath: path.resolve(snapshot.filePath), existed: false };
  if (!snapshotPath) throw new Error(`Office 事务 ${phase} 快照缺失: ${snapshot.filePath}`);
  const expectedRoot = transactionDirectory(root, id);
  const source = path.resolve(snapshotPath);
  const relative = path.relative(expectedRoot, source);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) throw new Error("Office 事务快照不在受控目录中");
  return { filePath: path.resolve(snapshot.filePath), existed: true, snapshotPath: source };
}

async function captureFileState(
  filePath: string,
  snapshotDir: string,
  index: number,
): Promise<{ existed: boolean; snapshotPath?: string; hash?: string; directory?: boolean }> {
  const resolved = path.resolve(filePath);
  try {
    const info = await stat(resolved);
    if (!info.isFile()) return { existed: false, directory: info.isDirectory() };
    const snapshotPath = path.join(snapshotDir, `${index}${path.extname(resolved) || ".bin"}`);
    await copyFile(resolved, snapshotPath);
    return { existed: true, snapshotPath, hash: await hashFile(resolved) };
  } catch (error) {
    if (isFileNotFound(error)) return { existed: false };
    throw error;
  }
}

async function detectTransactionConflicts(
  record: OfficeTransactionRecord,
  expected: "before" | "after",
): Promise<OfficeTransactionConflict[]> {
  const conflicts: OfficeTransactionConflict[] = [];
  for (const snapshot of record.snapshots) {
    const current = await currentFileState(snapshot.filePath);
    const expectedExists = expected === "before" ? snapshot.existed : snapshot.afterExisted === true;
    const expectedHash = expected === "before" ? snapshot.beforeHash : snapshot.afterHash;
    if (current.existed !== expectedExists) {
      conflicts.push({ filePath: snapshot.filePath, expected, reason: expectedExists ? "文件已被删除" : "出现了事务外文件" });
    } else if (expectedExists && expectedHash && current.hash !== expectedHash) {
      conflicts.push({ filePath: snapshot.filePath, expected, reason: "文件内容已在事务外修改" });
    }
  }
  return conflicts;
}

async function saveConflict(
  root: string,
  record: OfficeTransactionRecord,
  conflicts: OfficeTransactionConflict[],
  action: string,
): Promise<OfficeTransactionRecord> {
  const baseStatus = record.status === "conflicted" ? record.conflictBaseStatus : record.status;
  record.status = "conflicted";
  record.conflictBaseStatus = baseStatus;
  record.conflicts = conflicts;
  record.error = `${action}已拦截：${conflicts.length} 个文件在事务外发生变化`;
  await saveOfficeTransaction(root, record);
  return record;
}

async function currentFileState(filePath: string): Promise<{ existed: boolean; hash?: string }> {
  try {
    const info = await stat(filePath);
    return info.isFile() ? { existed: true, hash: await hashFile(filePath) } : { existed: false };
  } catch (error) {
    if (isFileNotFound(error)) return { existed: false };
    throw error;
  }
}

async function fileExists(filePath: string): Promise<boolean> {
  try { return (await stat(filePath)).isFile(); }
  catch (error) {
    if (isFileNotFound(error)) return false;
    throw error;
  }
}

async function hashFile(filePath: string): Promise<string> {
  return createHash("sha256").update(await readFile(filePath)).digest("hex");
}

function hasAfterState(record: OfficeTransactionRecord): boolean {
  return record.snapshots.every((snapshot) => snapshot.afterExisted !== undefined);
}

function addPath(target: Map<string, string>, value?: string): void {
  if (!value || !path.isAbsolute(value)) return;
  const resolved = path.resolve(value);
  target.set(process.platform === "win32" ? resolved.toLowerCase() : resolved, resolved);
}

function samePath(left: string, right: string): boolean {
  const a = path.resolve(left);
  const b = path.resolve(right);
  return process.platform === "win32" ? a.toLowerCase() === b.toLowerCase() : a === b;
}

function transactionDirectory(root: string, id: string): string {
  validateRecordId(id);
  return path.join(path.resolve(root), id);
}

function validateRecordId(id: string): void {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(id)) {
    throw new Error("Office 事务 ID 无效");
  }
}

function isFileNotFound(error: unknown): boolean {
  return Boolean(error && typeof error === "object" && "code" in error && error.code === "ENOENT");
}
