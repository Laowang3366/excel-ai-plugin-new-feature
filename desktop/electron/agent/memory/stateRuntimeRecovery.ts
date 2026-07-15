import * as fs from "fs";
import * as path from "path";

import { openSqliteDatabase, runPragmaValue } from "../storage/nodeSqlite";
import type { RuntimeDbName, RuntimeRecoveryReport } from "./stateRuntimeTypes";
import type { SqliteDatabase } from "../storage/nodeSqlite";

export function openRuntimeDatabaseWithRecovery(
  dbPath: string,
  dbName: RuntimeDbName,
  reports: RuntimeRecoveryReport[],
): SqliteDatabase {
  if (dbPath === ":memory:") return openSqliteDatabase(dbPath);

  let db: SqliteDatabase | undefined;
  try {
    db = openSqliteDatabase(dbPath);
    assertDatabaseHealthy(db);
    return db;
  } catch (error) {
    try {
      db?.close();
    } catch {
      // A corrupted database handle may already be unusable.
    }
    if (!isRecoverableSqliteCorruption(error)) {
      throw error;
    }
    const reason = error instanceof Error ? error.message : String(error);
    const backupPaths = backupAndRemoveDatabaseFiles(dbPath, dbName);
    reports.push({
      dbName,
      dbPath,
      reason,
      backupPaths,
      recoveredAt: Date.now(),
    });
    return openSqliteDatabase(dbPath);
  }
}

function assertDatabaseHealthy(db: SqliteDatabase): void {
  const result = runPragmaValue(db, "quick_check");
  if (String(result).toLowerCase() !== "ok") {
    throw new Error(`SQLite quick_check failed: ${String(result)}`);
  }
}

function backupAndRemoveDatabaseFiles(dbPath: string, dbName: RuntimeDbName): string[] {
  const existingPaths = [dbPath, `${dbPath}-wal`, `${dbPath}-shm`].filter((filePath) =>
    fs.existsSync(filePath),
  );
  if (existingPaths.length === 0) return [];

  const recoveryDir = path.join(path.dirname(dbPath), "recovery");
  fs.mkdirSync(recoveryDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const backupPaths: string[] = [];

  for (const sourcePath of existingPaths) {
    const suffix = sourcePath.slice(dbPath.length);
    const backupPath = path.join(recoveryDir, `${dbName}-${stamp}${suffix || ".db"}.bak`);
    fs.copyFileSync(sourcePath, backupPath);
    fs.rmSync(sourcePath, { force: true });
    backupPaths.push(backupPath);
  }

  return backupPaths;
}

function isRecoverableSqliteCorruption(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return /database disk image is malformed|file is not a database|malformed database schema|SQLite quick_check failed/i.test(
    message,
  );
}
