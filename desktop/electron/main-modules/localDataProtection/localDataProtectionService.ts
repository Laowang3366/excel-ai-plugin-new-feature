import { app } from "electron";
import { createLogger } from "../../shared/logger";
import { openOrCreateDataKeystore, type DataKeystore } from "./dataKeystore";
import {
  migrateLocalDataAtomically,
  readProtectionMarker,
  type AtomicMigrationHandle,
  type FinalizeResult,
} from "./localDataMigrator";
import { recoverInterruptedMigration as recoverInterruptedMigrationImpl } from "./localDataRecovery";
import {
  openOrCreateManagedReplicaRegistry,
  type ManagedReplicaRegistry,
} from "./managedReplicaRegistry";
import {
  createPayloadProtection,
  setPayloadProtection,
  type PayloadProtection,
} from "./payloadProtection";

const logger = createLogger("LocalDataProtection");

export interface LocalDataProtectionRuntime {
  keystore: DataKeystore;
  registry: ManagedReplicaRegistry;
  protection: PayloadProtection;
}

let runtime: LocalDataProtectionRuntime | null = null;

export function getLocalDataProtectionRuntime(): LocalDataProtectionRuntime | null {
  return runtime;
}

function refreshProtection(): void {
  if (!runtime) return;
  runtime.protection = createPayloadProtection(runtime.keystore);
  setPayloadProtection(runtime.protection);
}

function reportFinalizeFailure(result: FinalizeResult): void {
  if (result.cleared || !result.backupRoot) return;
  runtime?.registry.upsert({
    category: "old_root",
    absolutePath: result.backupRoot,
    status: "pending_erase",
    notes: result.error ?? "backup_cleanup_failed",
  });
  logger.warn("Encryption backup cleanup failed; registered pending erase", {
    backupRoot: result.backupRoot,
    error: result.error,
  });
}

export function recoverInterruptedMigration(options: {
  dataRoot: string;
  userDataPath?: string;
}): void {
  if (!runtime) return;
  recoverInterruptedMigrationImpl({
    ...options,
    runtime,
    refreshProtection,
    reportFinalizeFailure,
  });
}

export function initializeLocalDataProtection(options: {
  userDataPath?: string;
  dataRoot: string;
  cipher?: Parameters<typeof openOrCreateDataKeystore>[0] extends infer T
    ? T extends { cipher?: infer C }
      ? C
      : never
    : never;
}): LocalDataProtectionRuntime {
  const userDataPath = options.userDataPath ?? app.getPath("userData");
  // Gate is inside openOrCreateDataKeystore(dataRoot) — create and empty reseed share it.
  const opened = openOrCreateDataKeystore({
    userDataPath,
    cipher: options.cipher,
    dataRoot: options.dataRoot,
  });
  const registry = openOrCreateManagedReplicaRegistry({
    installId: opened.keystore.installId,
    userDataPath,
  });
  registry.upsert({
    category: "active_root",
    absolutePath: options.dataRoot,
    status: "active",
  });
  setPayloadProtection(createPayloadProtection(opened.keystore));
  runtime = {
    keystore: opened.keystore,
    registry,
    protection: createPayloadProtection(opened.keystore),
  };

  recoverInterruptedMigration({ dataRoot: options.dataRoot, userDataPath });
  refreshProtection();

  logger.info("Local data protection ready", {
    createdKeystore: opened.created,
    keyId: opened.keystore.currentKeyId,
  });
  return runtime;
}

export async function ensureLocalDataEncrypted(options: {
  dataRoot: string;
  seal: () => Promise<void>;
  restore: () => Promise<void>;
  userDataPath?: string;
}): Promise<{ migrated: boolean; contentKeyId: number; finalizeWarning?: string }> {
  if (!runtime) throw new Error("local_data_protection_uninitialized");
  const marker = readProtectionMarker(options.dataRoot);
  if (marker?.contentKeyId === runtime.keystore.currentKeyId) {
    return { migrated: false, contentKeyId: runtime.keystore.currentKeyId };
  }
  await options.seal();
  let handle: AtomicMigrationHandle | null = null;
  try {
    const result = await migrateLocalDataAtomically({
      dataRoot: options.dataRoot,
      protection: runtime.protection,
      targetKeyId: runtime.keystore.currentKeyId,
      previousKeyId: runtime.keystore.currentKeyId,
      kind: "encrypt",
      userDataPath: options.userDataPath,
    });
    if ("migrated" in result && result.migrated === false) {
      await options.restore();
      return { migrated: false, contentKeyId: result.contentKeyId };
    }
    handle = result as AtomicMigrationHandle;
    handle.markCommitted();
    await options.restore();
    const finalized = await handle.finalize();
    if (!finalized.cleared) {
      reportFinalizeFailure(finalized);
      return {
        migrated: true,
        contentKeyId: handle.targetKeyId,
        finalizeWarning: finalized.error,
      };
    }
    return { migrated: true, contentKeyId: handle.targetKeyId };
  } catch (error) {
    if (handle) {
      try {
        await handle.rollback();
      } catch (rollbackError) {
        throw new Error(
          `${error instanceof Error ? error.message : String(error)}; rollback failed: ${
            rollbackError instanceof Error ? rollbackError.message : String(rollbackError)
          }`,
        );
      }
    }
    await options.restore().catch(() => {});
    throw error;
  }
}

export async function rotateLocalDataKey(options: {
  dataRoot: string;
  seal: () => Promise<void>;
  restore: () => Promise<void>;
  userDataPath?: string;
  afterSwapBeforeCommit?: () => void | Promise<void>;
  afterCommitBeforeRestore?: () => void | Promise<void>;
}): Promise<{ success: boolean; keyId?: number; error?: string; finalizeWarning?: string }> {
  if (!runtime) return { success: false, error: "本地数据保护未初始化" };
  const previousKeyId = runtime.keystore.currentKeyId;
  let handle: AtomicMigrationHandle | null = null;
  let committed = false;
  try {
    await options.seal();
    const pendingKeyId = runtime.keystore.createPendingKey();
    const result = await migrateLocalDataAtomically({
      dataRoot: options.dataRoot,
      protection: runtime.protection,
      targetKeyId: pendingKeyId,
      previousKeyId,
      kind: "rotate",
      userDataPath: options.userDataPath,
      force: true,
    });
    if ("migrated" in result && result.migrated === false) {
      runtime.keystore.discardPendingKey();
      await options.restore();
      return { success: true, keyId: previousKeyId };
    }
    handle = result as AtomicMigrationHandle;
    if (options.afterSwapBeforeCommit) await options.afterSwapBeforeCommit();
    runtime.keystore.commitPendingKey();
    committed = true;
    handle.markCommitted();
    refreshProtection();
    if (options.afterCommitBeforeRestore) await options.afterCommitBeforeRestore();
    await options.restore();
    runtime.keystore.purgeRetiredKeys();
    const finalized = await handle.finalize();
    if (!finalized.cleared) {
      reportFinalizeFailure(finalized);
      return {
        success: true,
        keyId: runtime.keystore.currentKeyId,
        finalizeWarning: finalized.error,
      };
    }
    return { success: true, keyId: runtime.keystore.currentKeyId };
  } catch (error) {
    const parts = [error instanceof Error ? error.message : String(error)];
    try {
      if (handle) {
        try {
          await handle.rollback();
        } catch (rollbackError) {
          parts.push(
            `rollback failed: ${
              rollbackError instanceof Error ? rollbackError.message : String(rollbackError)
            }`,
          );
          // Keep keys that can still decrypt whatever remains on disk.
          refreshProtection();
          await options.restore().catch(() => {});
          return { success: false, keyId: previousKeyId, error: parts.join("; ") };
        }
      }
      if (committed) {
        runtime.keystore.revertToKeyId(previousKeyId);
      } else {
        runtime.keystore.discardPendingKey();
      }
      refreshProtection();
      await options.restore();
    } catch (recoveryError) {
      parts.push(
        `recovery failed: ${
          recoveryError instanceof Error ? recoveryError.message : String(recoveryError)
        }`,
      );
    }
    return { success: false, keyId: previousKeyId, error: parts.join("; ") };
  }
}

export function registerManagedDataRoot(
  absolutePath: string,
  category: "active_root" | "old_root" | "privacy_export",
): void {
  if (!runtime) return;
  runtime.registry.upsert({
    category,
    absolutePath,
    status: category === "old_root" ? "pending_erase" : "active",
  });
}

export function clearLocalDataProtectionRuntimeForTests(): void {
  runtime = null;
  setPayloadProtection(null);
}

export function isLocalDataProtectionCurrent(dataRoot: string): boolean {
  if (!runtime) return false;
  const marker = readProtectionMarker(dataRoot);
  return marker?.contentKeyId === runtime.keystore.currentKeyId;
}
