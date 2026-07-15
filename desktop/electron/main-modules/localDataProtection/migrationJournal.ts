import * as fs from "node:fs";
import * as path from "node:path";
import { app } from "electron";

export type MigrationPhase =
  "idle" | "staging" | "swapping" | "swapped" | "committed" | "finalizing";

export interface MigrationJournalRecord {
  formatVersion: 1;
  phase: MigrationPhase;
  kind: "encrypt" | "rotate";
  dataRoot: string;
  stageRoot: string;
  backupRoot: string;
  targetKeyId: number;
  previousKeyId: number;
  updatedAt: string;
}

const JOURNAL_FILE = "local-data-migration-journal.json";

export function migrationJournalPath(userDataPath = app.getPath("userData")): string {
  return path.join(userDataPath, JOURNAL_FILE);
}

/** Fixed sibling paths for a data root — no pid/time, recoverable after crash. */
export function fixedTransactionPaths(dataRoot: string): {
  stageRoot: string;
  backupRoot: string;
} {
  const resolved = path.resolve(dataRoot);
  const parent = path.dirname(resolved);
  const base = path.basename(resolved);
  return {
    stageRoot: path.join(parent, `.${base}.wengge-ldp-stage`),
    backupRoot: path.join(parent, `.${base}.wengge-ldp-backup`),
  };
}

export function readMigrationJournal(userDataPath?: string): MigrationJournalRecord | null {
  const filePath = migrationJournalPath(userDataPath);
  if (!fs.existsSync(filePath)) return null;
  const raw = JSON.parse(fs.readFileSync(filePath, "utf8")) as MigrationJournalRecord;
  if (raw.formatVersion !== 1 || !raw.dataRoot || !raw.phase) {
    throw new Error("invalid_migration_journal");
  }
  return raw;
}

export function writeMigrationJournal(record: MigrationJournalRecord, userDataPath?: string): void {
  const filePath = migrationJournalPath(userDataPath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const next: MigrationJournalRecord = {
    ...record,
    updatedAt: new Date().toISOString(),
  };
  const tempPath = `${filePath}.${process.pid}.tmp`;
  fs.writeFileSync(tempPath, `${JSON.stringify(next, null, 2)}\n`, "utf8");
  fs.renameSync(tempPath, filePath);
}

export function clearMigrationJournal(userDataPath?: string): void {
  const filePath = migrationJournalPath(userDataPath);
  if (fs.existsSync(filePath)) {
    fs.rmSync(filePath, { force: false });
  }
}
