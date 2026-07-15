import * as fs from "node:fs";
import { readProtectionMarker } from "./localDataMigrator";
import { fixedTransactionPaths, migrationJournalPath } from "./migrationJournal";

/**
 * True when local disk still holds protected data or any incomplete protection transaction
 * for this data root. Used to refuse empty-keystore reseed AND brand-new keystore creation.
 *
 * Journal path trust is NOT used here: any journal file blocks minting keys. Recovery alone
 * decides whether journal paths are trusted enough to follow (fixedTransactionPaths only).
 */
export function hasProtectedLocalDataArtifacts(dataRoot: string, userDataPath: string): boolean {
  if (readProtectionMarker(dataRoot) != null) return true;

  const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
  if (readProtectionMarker(stageRoot) != null) return true;
  if (readProtectionMarker(backupRoot) != null) return true;
  // Incomplete rename window: stage/backup dirs may exist without a readable marker yet.
  if (fs.existsSync(stageRoot) || fs.existsSync(backupRoot)) return true;

  // Any migration journal blocks key minting until recovery handles it — including
  // valid-but-untrusted/tampered path journals (never treat as "disk clean").
  if (fs.existsSync(migrationJournalPath(userDataPath))) return true;

  return false;
}
