import { access, mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { runLocalDataMaintenance } from "./localDataMaintenance";

const NOW = Date.parse("2026-07-15T12:00:00.000Z");
const OLD = "2026-05-01T00:00:00.000Z";
const VERY_OLD = "2026-03-01T00:00:00.000Z";
const RECENT = "2026-07-10T00:00:00.000Z";

describe("local data maintenance", () => {
  it("prunes expired managed data while preserving active records and unrelated files", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "local-data-maintenance-"));
    try {
      const oldLog = path.join(root, "logs", "app-2026-05-01.log");
      const currentLog = path.join(root, "logs", "app-2026-07-15.log");
      const unrelatedLog = path.join(root, "logs", "operator-notes.txt");
      await writeFiles([
        [oldLog, "old log"],
        [currentLog, "current log"],
        [unrelatedLog, "keep"],
      ]);

      const backupRoot = path.join(root, "office-backups");
      const oldBackup = await writeBackup(backupRoot, "old-backup", OLD);
      const recentBackup = await writeBackup(backupRoot, "recent-backup", RECENT);

      const transactionRoot = path.join(root, "office-automation", "transactions");
      const oldTransaction = await writeTransaction(
        transactionRoot,
        "11111111-1111-4111-8111-111111111111",
        "applied",
        OLD,
      );
      const pendingTransaction = await writeTransaction(
        transactionRoot,
        "22222222-2222-4222-8222-222222222222",
        "pending",
        OLD,
      );

      const workflowRoot = path.join(root, "office-automation", "workflows");
      const oldWorkflow = await writeWorkflow(
        workflowRoot,
        "33333333-3333-4333-8333-333333333333",
        "done",
        VERY_OLD,
      );
      const runningWorkflow = await writeWorkflow(
        workflowRoot,
        "44444444-4444-4444-8444-444444444444",
        "running",
        OLD,
      );
      const template = path.join(workflowRoot, "templates", "55555555-5555-4555-8555-555555555555.json");
      await writeFiles([[template, "{}"]]);

      const report = await runLocalDataMaintenance(root, { now: NOW });

      await expectMissing(oldLog, oldBackup.dataPath, oldBackup.metadataPath, oldTransaction, oldWorkflow);
      await expectPresent(
        currentLog,
        unrelatedLog,
        recentBackup.dataPath,
        recentBackup.metadataPath,
        pendingTransaction,
        runningWorkflow,
        template,
      );
      expect(report.deletedFiles).toBeGreaterThanOrEqual(4);
      expect(report.deletedDirectories).toBe(1);
      expect(report.reclaimedBytes).toBeGreaterThan(0);
      expect(report.errors).toEqual([]);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it("applies quotas without deleting the current log or paused workflow", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "local-data-maintenance-quota-"));
    try {
      const currentLog = path.join(root, "logs", "app-2026-07-15.log");
      const previousLog = path.join(root, "logs", "app-2026-07-14.log");
      await writeFiles([[currentLog, "current"], [previousLog, "previous"]]);

      const workflowRoot = path.join(root, "office-automation", "workflows");
      const pausedWorkflow = await writeWorkflow(
        workflowRoot,
        "66666666-6666-4666-8666-666666666666",
        "paused",
        RECENT,
      );
      const doneWorkflow = await writeWorkflow(
        workflowRoot,
        "77777777-7777-4777-8777-777777777777",
        "done",
        RECENT,
      );

      await runLocalDataMaintenance(root, {
        now: NOW,
        policy: {
          logs: { maxAgeDays: 30, maxEntries: 1, maxBytes: 1024 },
          workflows: { maxAgeDays: 90, maxEntries: 0, maxBytes: 0 },
        },
      });

      await expectPresent(currentLog, pausedWorkflow);
      await expectMissing(previousLog, doneWorkflow);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it("reports one category failure and continues maintaining the others", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "local-data-maintenance-error-"));
    try {
      await writeFile(path.join(root, "logs"), "not a directory");
      const oldBackup = await writeBackup(path.join(root, "office-backups"), "old-backup", OLD);

      const report = await runLocalDataMaintenance(root, { now: NOW });

      expect(report.errors.some((error) => error.startsWith("logs:"))).toBe(true);
      await expectMissing(oldBackup.dataPath, oldBackup.metadataPath);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

async function writeBackup(root: string, id: string, createdAt: string) {
  await mkdir(root, { recursive: true });
  const dataPath = path.join(root, `${id}.xlsx`);
  const metadataPath = path.join(root, `${id}.json`);
  await writeFile(dataPath, "backup");
  await writeFile(metadataPath, `${JSON.stringify({
    id,
    app: "excel",
    operation: "writeRange",
    sourcePath: path.join(root, "source.xlsx"),
    backupPath: dataPath,
    createdAt,
    size: 6,
  })}\n`);
  return { dataPath, metadataPath };
}

async function writeTransaction(root: string, id: string, status: string, updatedAt: string) {
  const directory = path.join(root, id);
  await mkdir(directory, { recursive: true });
  await writeFile(path.join(directory, "transaction.json"), `${JSON.stringify({ id, status, updatedAt })}\n`);
  await writeFile(path.join(directory, "snapshot.bin"), "snapshot");
  return directory;
}

async function writeWorkflow(root: string, id: string, status: string, updatedAt: string) {
  await mkdir(root, { recursive: true });
  const filePath = path.join(root, `${id}.json`);
  await writeFile(filePath, `${JSON.stringify({
    id,
    status,
    updatedAt,
    ...(status === "running" ? { leaseExpiresAt: "2026-07-16T00:00:00.000Z" } : {}),
  })}\n`);
  return filePath;
}

async function writeFiles(files: Array<[string, string]>): Promise<void> {
  for (const [filePath, content] of files) {
    await mkdir(path.dirname(filePath), { recursive: true });
    await writeFile(filePath, content);
  }
}

async function expectMissing(...paths: string[]): Promise<void> {
  for (const filePath of paths) {
    await expect(access(filePath)).rejects.toMatchObject({ code: "ENOENT" });
  }
}

async function expectPresent(...paths: string[]): Promise<void> {
  for (const filePath of paths) {
    await expect(access(filePath)).resolves.toBeUndefined();
  }
}
