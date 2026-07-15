import * as fs from "node:fs";
import * as path from "node:path";
import { app } from "electron";
import { createLogger } from "../../shared/logger";
import type { FinalizeResult } from "./localDataMigrator";
import { readProtectionMarker } from "./localDataMigrator";
import {
  clearMigrationJournal,
  fixedTransactionPaths,
  readMigrationJournal,
  writeMigrationJournal,
} from "./migrationJournal";
import type { LocalDataProtectionRuntime } from "./localDataProtectionService";

const logger = createLogger("LocalDataRecovery");

function pathsEqual(left: string, right: string): boolean {
  return (
    path
      .resolve(left)
      .replace(/[\\/]+$/u, "")
      .toLowerCase() ===
    path
      .resolve(right)
      .replace(/[\\/]+$/u, "")
      .toLowerCase()
  );
}

function assertJournalPathsTrusted(
  journal: NonNullable<ReturnType<typeof readMigrationJournal>>,
  knownDataRoot: string,
): void {
  const expected = fixedTransactionPaths(knownDataRoot);
  if (!pathsEqual(journal.dataRoot, knownDataRoot)) {
    throw new Error("migration_journal_data_root_mismatch");
  }
  if (!pathsEqual(journal.stageRoot, expected.stageRoot)) {
    throw new Error("migration_journal_stage_path_mismatch");
  }
  if (!pathsEqual(journal.backupRoot, expected.backupRoot)) {
    throw new Error("migration_journal_backup_path_mismatch");
  }
}

export function recoverInterruptedMigration(options: {
  dataRoot: string;
  userDataPath?: string;
  runtime: LocalDataProtectionRuntime;
  refreshProtection: () => void;
  reportFinalizeFailure: (result: FinalizeResult) => void;
}): void {
  const userDataPath = options.userDataPath ?? app.getPath("userData");
  const journal = readMigrationJournal(userDataPath);
  if (!journal) return;

  assertJournalPathsTrusted(journal, options.dataRoot);

  const dataRoot = path.resolve(options.dataRoot);
  const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
  const targetKeyId = journal.targetKeyId;
  const dataExists = fs.existsSync(dataRoot);
  const backupExists = fs.existsSync(backupRoot);
  const stageExists = fs.existsSync(stageRoot);
  const marker = dataExists ? readProtectionMarker(dataRoot) : null;
  const { runtime, refreshProtection, reportFinalizeFailure } = options;

  const ensureTargetKeyReady = (): void => {
    if (runtime.keystore.currentKeyId === targetKeyId) {
      // Commit may have succeeded before crash; still drop retired prior keys.
      runtime.keystore.purgeRetiredKeys();
      refreshProtection();
      return;
    }
    if (runtime.keystore.pendingKeyId === targetKeyId) {
      runtime.keystore.commitPendingKey();
      runtime.keystore.purgeRetiredKeys();
      refreshProtection();
      return;
    }
    throw new Error("migration_recovery_target_key_unavailable");
  };

  const cleanupBackupAfterKeyReady = (): void => {
    ensureTargetKeyReady();
    if (!fs.existsSync(backupRoot)) {
      clearMigrationJournal(userDataPath);
      return;
    }
    try {
      fs.rmSync(backupRoot, { recursive: true, force: false });
      clearMigrationJournal(userDataPath);
    } catch (error) {
      reportFinalizeFailure({
        cleared: false,
        backupRoot,
        error: error instanceof Error ? error.message : String(error),
      });
      writeMigrationJournal({ ...journal, phase: "committed" }, userDataPath);
    }
  };

  if (journal.phase === "staging") {
    if (stageExists) fs.rmSync(stageRoot, { recursive: true, force: false });
    if (runtime.keystore.pendingKeyId != null) runtime.keystore.discardPendingKey();
    clearMigrationJournal(userDataPath);
    refreshProtection();
    logger.info("Recovered staging phase: discarded stage/pending");
    return;
  }

  if (journal.phase === "swapping") {
    if (!dataExists && backupExists && stageExists) {
      fs.renameSync(backupRoot, dataRoot);
      fs.rmSync(stageRoot, { recursive: true, force: false });
      if (runtime.keystore.pendingKeyId != null) runtime.keystore.discardPendingKey();
      clearMigrationJournal(userDataPath);
      refreshProtection();
      logger.info("Recovered swapping rename-gap: restored backup");
      return;
    }
    if (dataExists && backupExists && !stageExists && marker?.contentKeyId === targetKeyId) {
      writeMigrationJournal({ ...journal, phase: "swapped" }, userDataPath);
      cleanupBackupAfterKeyReady();
      logger.info("Recovered swapping post-swap window as swapped");
      return;
    }
    if (dataExists && !backupExists && stageExists) {
      fs.rmSync(stageRoot, { recursive: true, force: false });
      if (runtime.keystore.pendingKeyId != null) runtime.keystore.discardPendingKey();
      clearMigrationJournal(userDataPath);
      refreshProtection();
      logger.info("Recovered swapping pre-rename: discarded stage/pending");
      return;
    }
    throw new Error(
      `migration_recovery_swapping_inconsistent:data=${dataExists},backup=${backupExists},stage=${stageExists},marker=${marker?.contentKeyId ?? "none"}`,
    );
  }

  if (
    journal.phase === "swapped" ||
    journal.phase === "committed" ||
    journal.phase === "finalizing"
  ) {
    cleanupBackupAfterKeyReady();
    logger.info("Recovered post-swap phase: ensured target key and cleaned backup");
    return;
  }

  throw new Error(`migration_recovery_unknown_phase:${journal.phase}`);
}
