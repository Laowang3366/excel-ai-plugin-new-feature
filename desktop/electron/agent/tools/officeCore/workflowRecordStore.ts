import { randomUUID } from "node:crypto";
import { mkdir, open, readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";

import type { OfficeWorkflowRecord } from "./workflow";

interface WorkflowLockRecord {
  workflowId: string;
  token: string;
  processId: number;
  expiresAt: string;
}

export interface OfficeWorkflowLock {
  workflowId: string;
  renew(): Promise<void>;
  assertOwned(): Promise<void>;
  release(): Promise<void>;
}

const ACTIVE_WORKFLOW_LOCKS = new Set<string>();
const WORKFLOW_LOCK_LEASE_MS = 120_000;

export async function getOfficeWorkflowRecord(root: string, id: string): Promise<OfficeWorkflowRecord> {
  validateWorkflowId(id);
  const record = JSON.parse(await readFile(workflowPath(root, id), "utf8")) as OfficeWorkflowRecord;
  if (record.id !== id || !Array.isArray(record.steps) || !Array.isArray(record.stepRecords)) {
    throw new Error("Office 工作流记录已损坏");
  }
  return record;
}

export async function listOfficeWorkflowRecords(root: string): Promise<OfficeWorkflowRecord[]> {
  let names: string[];
  try { names = await readdir(path.resolve(root)); } catch { return []; }
  const records = await Promise.all(names.filter((name) => name.endsWith(".json")).map(async (name) => {
    try { return await getOfficeWorkflowRecord(root, path.basename(name, ".json")); } catch { return undefined; }
  }));
  return records
    .filter((record): record is OfficeWorkflowRecord => Boolean(record))
    .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
}

export async function saveOfficeWorkflowRecord(root: string, record: OfficeWorkflowRecord): Promise<void> {
  record.updatedAt = new Date().toISOString();
  await mkdir(path.resolve(root), { recursive: true });
  const destination = workflowPath(root, record.id);
  const temporary = `${destination}.${randomUUID()}.tmp`;
  await writeFile(temporary, `${JSON.stringify(record, null, 2)}\n`, "utf8");
  await rm(destination, { force: true });
  await rename(temporary, destination);
}

export async function requestOfficeWorkflowCancellation(root: string, id: string): Promise<OfficeWorkflowRecord> {
  const record = await getOfficeWorkflowRecord(root, id);
  if (["done", "failed", "cancelled"].includes(record.status)) return record;
  record.cancelRequested = true;
  await saveOfficeWorkflowRecord(root, record);
  return record;
}

export async function isOfficeWorkflowCancellationRequested(
  root: string | undefined,
  inMemoryRecord: OfficeWorkflowRecord,
): Promise<boolean> {
  if (inMemoryRecord.cancelRequested) return true;
  if (!root) return false;
  try { return (await getOfficeWorkflowRecord(root, inMemoryRecord.id)).cancelRequested === true; }
  catch { return false; }
}

export async function acquireOfficeWorkflowLock(root: string, workflowId: string): Promise<OfficeWorkflowLock> {
  validateWorkflowId(workflowId);
  const key = `${path.resolve(root).toLowerCase()}|${workflowId}`;
  if (ACTIVE_WORKFLOW_LOCKS.has(key)) throw new Error(`Office 工作流 ${workflowId} 正在运行`);
  const directory = path.join(path.resolve(root), ".locks");
  const lockPath = path.join(directory, `${workflowId}.lock`);
  await mkdir(directory, { recursive: true });
  const token = randomUUID();

  for (let attempt = 0; attempt < 4; attempt++) {
    const record = newLockRecord(workflowId, token);
    try {
      const handle = await open(lockPath, "wx");
      try { await handle.writeFile(`${JSON.stringify(record)}\n`, "utf8"); }
      finally { await handle.close(); }
      ACTIVE_WORKFLOW_LOCKS.add(key);
      return createWorkflowLock(lockPath, key, record);
    } catch (error) {
      if (!isAlreadyExists(error)) throw error;
      if (!(await existingLockIsStale(lockPath))) throw new Error(`Office 工作流 ${workflowId} 正在运行`);
      await rm(lockPath, { force: true });
    }
  }
  throw new Error(`无法取得 Office 工作流 ${workflowId} 的执行锁`);
}

export function startOfficeWorkflowLockHeartbeat(lock: OfficeWorkflowLock): () => void {
  const timer = setInterval(() => { void lock.renew().catch(() => undefined); }, 20_000);
  timer.unref?.();
  return () => clearInterval(timer);
}

function createWorkflowLock(lockPath: string, key: string, initial: WorkflowLockRecord): OfficeWorkflowLock {
  let record = initial;
  let released = false;
  const assertOwned = async () => {
    if (released) throw new Error(`Office 工作流 ${record.workflowId} 的执行锁已释放`);
    const current = await readWorkflowLock(lockPath);
    if (!current || current.token !== record.token || current.processId !== process.pid) {
      throw new Error(`Office 工作流 ${record.workflowId} 的执行锁已丢失`);
    }
  };
  return {
    workflowId: record.workflowId,
    assertOwned,
    renew: async () => {
      await assertOwned();
      record = newLockRecord(record.workflowId, record.token);
      await writeFile(lockPath, `${JSON.stringify(record)}\n`, "utf8");
    },
    release: async () => {
      if (released) return;
      released = true;
      try {
        const current = await readWorkflowLock(lockPath);
        if (current?.token === record.token) await rm(lockPath, { force: true });
      } finally {
        ACTIVE_WORKFLOW_LOCKS.delete(key);
      }
    },
  };
}

function newLockRecord(workflowId: string, token: string): WorkflowLockRecord {
  return {
    workflowId,
    token,
    processId: process.pid,
    expiresAt: new Date(Date.now() + WORKFLOW_LOCK_LEASE_MS).toISOString(),
  };
}

async function existingLockIsStale(lockPath: string): Promise<boolean> {
  const record = await readWorkflowLock(lockPath);
  if (record) return !isProcessRunning(record.processId);
  try { return Date.now() - (await stat(lockPath)).mtimeMs > WORKFLOW_LOCK_LEASE_MS; }
  catch { return true; }
}

async function readWorkflowLock(lockPath: string): Promise<WorkflowLockRecord | undefined> {
  try {
    const value = JSON.parse(await readFile(lockPath, "utf8")) as Partial<WorkflowLockRecord>;
    if (typeof value.workflowId !== "string" || typeof value.token !== "string" || !Number.isSafeInteger(value.processId)) return undefined;
    return value as WorkflowLockRecord;
  } catch { return undefined; }
}

function isProcessRunning(processId: number): boolean {
  if (!Number.isSafeInteger(processId) || processId <= 0) return false;
  try { process.kill(processId, 0); return true; }
  catch { return false; }
}

function isAlreadyExists(error: unknown): boolean {
  return Boolean(error && typeof error === "object" && "code" in error && error.code === "EEXIST");
}

function workflowPath(root: string, id: string): string {
  validateWorkflowId(id);
  return path.join(path.resolve(root), `${id}.json`);
}

function validateWorkflowId(id: string): void {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(id)) {
    throw new Error("Office 工作流 ID 无效");
  }
}
