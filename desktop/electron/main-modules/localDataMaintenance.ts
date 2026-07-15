import { readFile, readdir, rm, stat } from "node:fs/promises";
import path from "node:path";

import { pruneOfficeBackups } from "../agent/tools/officeCore/transactions";

const DAY_MS = 24 * 60 * 60 * 1000;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const TRANSACTION_STATUSES = new Set(["pending", "applied", "undone", "failed", "conflicted"]);
const WORKFLOW_STATUSES = new Set(["running", "paused", "done", "failed", "cancelled"]);

export interface RetentionLimit {
  maxAgeDays: number;
  maxEntries: number;
  maxBytes: number;
}

export interface LocalDataRetentionPolicy {
  logs: RetentionLimit;
  backups: RetentionLimit & { maxPerSource: number };
  transactions: RetentionLimit;
  workflows: RetentionLimit;
}

export interface LocalDataMaintenanceReport {
  deletedFiles: number;
  deletedDirectories: number;
  reclaimedBytes: number;
  errors: string[];
}

export const DEFAULT_LOCAL_DATA_RETENTION: LocalDataRetentionPolicy = {
  logs: { maxAgeDays: 30, maxEntries: 30, maxBytes: 100 * 1024 * 1024 },
  backups: {
    maxAgeDays: 30,
    maxEntries: 500,
    maxBytes: 2 * 1024 * 1024 * 1024,
    maxPerSource: 50,
  },
  transactions: { maxAgeDays: 30, maxEntries: 200, maxBytes: 2 * 1024 * 1024 * 1024 },
  workflows: { maxAgeDays: 90, maxEntries: 500, maxBytes: 100 * 1024 * 1024 },
};

interface ManagedEntry {
  entryPath: string;
  updatedAt: number;
  size: number;
  kind: "file" | "directory";
  protectedFromQuota?: boolean;
  protectedFromExpiry?: boolean;
}

export async function runLocalDataMaintenance(
  dataPath: string,
  options: {
    now?: number;
    policy?: Partial<LocalDataRetentionPolicy>;
  } = {},
): Promise<LocalDataMaintenanceReport> {
  const root = path.resolve(dataPath);
  const now = options.now ?? Date.now();
  const policy = mergePolicy(options.policy);
  const report = emptyReport();

  await captureMaintenanceErrors(report, "logs", async () => {
    const entries = await collectLogEntries(path.join(root, "logs"), now);
    mergeReport(report, await pruneManagedEntries(path.join(root, "logs"), entries, policy.logs, now));
  });

  await captureMaintenanceErrors(report, "office-backups", async () => {
    const backupResult = await pruneOfficeBackups(path.join(root, "office-backups"), {
      maxAgeDays: policy.backups.maxAgeDays,
      maxEntries: policy.backups.maxEntries,
      maxBytes: policy.backups.maxBytes,
      maxPerSource: policy.backups.maxPerSource,
      now,
    });
    report.deletedFiles += backupResult.deletedFiles;
    report.reclaimedBytes += backupResult.reclaimedBytes;
    report.errors.push(...backupResult.errors.map((error) => `office-backups: ${error}`));
  });

  await captureMaintenanceErrors(report, "office-transactions", async () => {
    const transactionRoot = path.join(root, "office-automation", "transactions");
    const entries = await collectTransactionEntries(transactionRoot, now);
    mergeReport(report, await pruneManagedEntries(transactionRoot, entries, policy.transactions, now));
  });

  await captureMaintenanceErrors(report, "office-workflows", async () => {
    const workflowRoot = path.join(root, "office-automation", "workflows");
    const entries = await collectWorkflowEntries(workflowRoot, now);
    mergeReport(report, await pruneManagedEntries(workflowRoot, entries, policy.workflows, now));
  });

  return report;
}

export function startLocalDataMaintenance(options: {
  getDataPath: () => string;
  intervalMs?: number;
  onReport?: (report: LocalDataMaintenanceReport) => void;
}): () => void {
  const intervalMs = options.intervalMs ?? 6 * 60 * 60 * 1000;
  let running = false;
  const timer = setInterval(() => {
    if (running) return;
    running = true;
    void runLocalDataMaintenance(options.getDataPath())
      .then((report) => options.onReport?.(report))
      .catch((error) => options.onReport?.({
        ...emptyReport(),
        errors: [errorMessage(error)],
      }))
      .finally(() => {
        running = false;
      });
  }, intervalMs);
  timer.unref?.();
  return () => clearInterval(timer);
}

async function collectLogEntries(root: string, now: number): Promise<ManagedEntry[]> {
  const entries = await readDirectory(root);
  const today = new Date(now).toISOString().slice(0, 10);
  const result: ManagedEntry[] = [];
  for (const entry of entries) {
    const match = /^app-(\d{4}-\d{2}-\d{2})\.log$/u.exec(entry.name);
    if (!entry.isFile() || !match) continue;
    const entryPath = path.join(root, entry.name);
    const updatedAt = Date.parse(`${match[1]}T00:00:00.000Z`);
    if (!Number.isFinite(updatedAt)) continue;
    const info = await stat(entryPath);
    result.push({
      entryPath,
      updatedAt,
      size: info.size,
      kind: "file",
      protectedFromExpiry: match[1] === today,
      protectedFromQuota: match[1] === today,
    });
  }
  return result;
}

async function collectTransactionEntries(root: string, now: number): Promise<ManagedEntry[]> {
  const entries = await readDirectory(root);
  const result: ManagedEntry[] = [];
  for (const entry of entries) {
    if (!entry.isDirectory() || !UUID_PATTERN.test(entry.name)) continue;
    const entryPath = path.join(root, entry.name);
    try {
      const record = JSON.parse(await readFile(path.join(entryPath, "transaction.json"), "utf8")) as {
        id?: string;
        status?: string;
        updatedAt?: string;
      };
      const updatedAt = Date.parse(record.updatedAt || "");
      if (
        record.id !== entry.name
        || !TRANSACTION_STATUSES.has(record.status || "")
        || !Number.isFinite(updatedAt)
      ) continue;
      const recentPending = record.status === "pending" && updatedAt >= now - 90 * DAY_MS;
      result.push({
        entryPath,
        updatedAt,
        size: await directorySize(entryPath),
        kind: "directory",
        protectedFromExpiry: recentPending,
        protectedFromQuota: record.status === "pending" || record.status === "conflicted",
      });
    } catch {
      // Corrupt records are preserved for manual recovery instead of being guessed safe to delete.
    }
  }
  return result;
}

async function collectWorkflowEntries(root: string, now: number): Promise<ManagedEntry[]> {
  const entries = await readDirectory(root);
  const result: ManagedEntry[] = [];
  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith(".json")) continue;
    const id = path.basename(entry.name, ".json");
    if (!UUID_PATTERN.test(id)) continue;
    const entryPath = path.join(root, entry.name);
    try {
      const record = JSON.parse(await readFile(entryPath, "utf8")) as {
        id?: string;
        status?: string;
        updatedAt?: string;
        leaseExpiresAt?: string;
      };
      const updatedAt = Date.parse(record.updatedAt || "");
      if (
        record.id !== id
        || !WORKFLOW_STATUSES.has(record.status || "")
        || !Number.isFinite(updatedAt)
      ) continue;
      const activeRunning = record.status === "running"
        && Date.parse(record.leaseExpiresAt || "") > now;
      const info = await stat(entryPath);
      result.push({
        entryPath,
        updatedAt,
        size: info.size,
        kind: "file",
        protectedFromExpiry: activeRunning,
        protectedFromQuota: activeRunning || record.status === "paused",
      });
    } catch {
      // Corrupt records are preserved for manual recovery instead of being guessed safe to delete.
    }
  }
  return result;
}

async function pruneManagedEntries(
  root: string,
  entries: ManagedEntry[],
  limits: RetentionLimit,
  now: number,
): Promise<LocalDataMaintenanceReport> {
  const report = emptyReport();
  const cutoff = now - limits.maxAgeDays * DAY_MS;
  let keptEntries = 0;
  let keptBytes = 0;
  const sorted = [...entries].sort((left, right) => right.updatedAt - left.updatedAt);

  for (const entry of sorted) {
    const expired = entry.updatedAt < cutoff && !entry.protectedFromExpiry;
    const overQuota = (
      keptEntries >= limits.maxEntries
      || keptBytes + entry.size > limits.maxBytes
    ) && !entry.protectedFromQuota;
    if (!expired && !overQuota) {
      keptEntries++;
      keptBytes += entry.size;
      continue;
    }

    try {
      assertManagedChild(root, entry.entryPath);
      await rm(entry.entryPath, { recursive: entry.kind === "directory", force: false });
      if (entry.kind === "directory") report.deletedDirectories++;
      else report.deletedFiles++;
      report.reclaimedBytes += entry.size;
    } catch (error) {
      report.errors.push(`${path.basename(entry.entryPath)}: ${errorMessage(error)}`);
    }
  }
  return report;
}

async function directorySize(root: string): Promise<number> {
  let total = 0;
  for (const entry of await readDirectory(root)) {
    const entryPath = path.join(root, entry.name);
    if (entry.isDirectory()) total += await directorySize(entryPath);
    else if (entry.isFile()) total += (await stat(entryPath)).size;
  }
  return total;
}

async function readDirectory(root: string) {
  try {
    return await readdir(root, { withFileTypes: true });
  } catch (error) {
    if (isFileNotFound(error)) return [];
    throw error;
  }
}

function assertManagedChild(root: string, candidate: string): void {
  const resolvedRoot = path.resolve(root);
  const resolvedCandidate = path.resolve(candidate);
  const relative = path.relative(resolvedRoot, resolvedCandidate);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`refusing to delete outside managed root: ${resolvedCandidate}`);
  }
}

function mergePolicy(overrides?: Partial<LocalDataRetentionPolicy>): LocalDataRetentionPolicy {
  return {
    logs: { ...DEFAULT_LOCAL_DATA_RETENTION.logs, ...overrides?.logs },
    backups: { ...DEFAULT_LOCAL_DATA_RETENTION.backups, ...overrides?.backups },
    transactions: { ...DEFAULT_LOCAL_DATA_RETENTION.transactions, ...overrides?.transactions },
    workflows: { ...DEFAULT_LOCAL_DATA_RETENTION.workflows, ...overrides?.workflows },
  };
}

async function captureMaintenanceErrors(
  report: LocalDataMaintenanceReport,
  category: string,
  operation: () => Promise<void>,
): Promise<void> {
  try {
    await operation();
  } catch (error) {
    report.errors.push(`${category}: ${errorMessage(error)}`);
  }
}

function emptyReport(): LocalDataMaintenanceReport {
  return { deletedFiles: 0, deletedDirectories: 0, reclaimedBytes: 0, errors: [] };
}

function mergeReport(target: LocalDataMaintenanceReport, source: LocalDataMaintenanceReport): void {
  target.deletedFiles += source.deletedFiles;
  target.deletedDirectories += source.deletedDirectories;
  target.reclaimedBytes += source.reclaimedBytes;
  target.errors.push(...source.errors);
}

function isFileNotFound(error: unknown): boolean {
  return Boolean(error && typeof error === "object" && "code" in error && error.code === "ENOENT");
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
