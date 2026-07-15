import { access, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { createOfficeBackup, listOfficeBackups, pruneOfficeBackups, restoreOfficeBackup } from "./transactions";

describe("Office file transactions", () => {
  it("creates, lists, and restores a persistent backup", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const backupRoot = path.join(tempDir, "backups");
      const sourcePath = path.join(tempDir, "book.xlsx");
      await writeFile(sourcePath, "before", "utf8");

      const record = await createOfficeBackup({
        backupRoot,
        app: "excel",
        operation: "styleTable",
        sourcePath,
      });
      await writeFile(sourcePath, "after", "utf8");

      expect(await listOfficeBackups(backupRoot, sourcePath)).toEqual([record]);
      await restoreOfficeBackup({ backupRoot, backupPath: record.backupPath, destinationPath: sourcePath });
      expect(await readFile(sourcePath, "utf8")).toBe("before");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("rejects backup paths outside the controlled directory", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const sourcePath = path.join(tempDir, "outside.docx");
      await writeFile(sourcePath, "outside", "utf8");
      await expect(restoreOfficeBackup({
        backupRoot: path.join(tempDir, "backups"),
        backupPath: sourcePath,
        destinationPath: path.join(tempDir, "target.docx"),
      })).rejects.toThrow("受控 Office 事务目录");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("rejects restoring a valid backup over a different Office file", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const backupRoot = path.join(tempDir, "backups");
      const sourcePath = path.join(tempDir, "source.xlsx");
      const otherPath = path.join(tempDir, "other.xlsx");
      await writeFile(sourcePath, "source", "utf8");
      await writeFile(otherPath, "other", "utf8");
      const record = await createOfficeBackup({
        backupRoot,
        app: "excel",
        operation: "styleTable",
        sourcePath,
      });

      await expect(restoreOfficeBackup({
        backupRoot,
        backupPath: record.backupPath,
        destinationPath: otherPath,
      })).rejects.toThrow("备份与目标文件不匹配");
      expect(await readFile(otherPath, "utf8")).toBe("other");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("prunes old backups without removing the newest recoverable versions", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const backupRoot = path.join(tempDir, "backups");
      const sourcePath = path.join(tempDir, "book.xlsx");
      await writeFile(sourcePath, "v1", "utf8");
      await createOfficeBackup({ backupRoot, app: "excel", operation: "edit1", sourcePath });
      await writeFile(sourcePath, "v2", "utf8");
      await createOfficeBackup({ backupRoot, app: "excel", operation: "edit2", sourcePath });
      await writeFile(sourcePath, "v3", "utf8");
      await createOfficeBackup({ backupRoot, app: "excel", operation: "edit3", sourcePath });

      await pruneOfficeBackups(backupRoot, { maxEntries: 2, maxPerSource: 2 });

      const records = await listOfficeBackups(backupRoot, sourcePath);
      expect(records).toHaveLength(2);
      expect(records.map((record) => record.operation)).toEqual(["edit3", "edit2"]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("prunes expired backups using the record timestamp", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const backupRoot = path.join(tempDir, "backups");
      const sourcePath = path.join(tempDir, "book.xlsx");
      await writeFile(sourcePath, "before", "utf8");
      const record = await createOfficeBackup({
        backupRoot,
        app: "excel",
        operation: "styleTable",
        sourcePath,
      });

      const result = await pruneOfficeBackups(backupRoot, {
        maxAgeDays: 30,
        now: Date.parse(record.createdAt) + 31 * 24 * 60 * 60 * 1000,
      });

      expect(result.deletedRecords).toBe(1);
      await expect(access(record.backupPath)).rejects.toMatchObject({ code: "ENOENT" });
      expect(await listOfficeBackups(backupRoot)).toEqual([]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("does not follow forged backup metadata outside the managed directory", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-transactions-"));
    try {
      const backupRoot = path.join(tempDir, "backups");
      const outsidePath = path.join(tempDir, "outside.xlsx");
      await mkdir(backupRoot, { recursive: true });
      await writeFile(outsidePath, "outside");
      await writeFile(path.join(backupRoot, "forged.json"), `${JSON.stringify({
        id: "forged",
        app: "excel",
        operation: "writeRange",
        sourcePath: path.join(tempDir, "source.xlsx"),
        backupPath: outsidePath,
        createdAt: "2020-01-01T00:00:00.000Z",
        size: 7,
      })}\n`);

      expect(await listOfficeBackups(backupRoot)).toEqual([]);
      await pruneOfficeBackups(backupRoot, { maxAgeDays: 1, now: Date.now() });
      await expect(access(outsidePath)).resolves.toBeUndefined();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
