import { createLogger } from "../../shared/logger";
import {
  ensureLocalDataEncrypted,
  initializeLocalDataProtection,
  registerManagedDataRoot,
  rotateLocalDataKey,
} from "./localDataProtectionService";
import { clearManagedDataRoot } from "./managedRootCleanup";

const logger = createLogger("LocalDataLifecycle");

export async function bootstrapLocalDataProtection(options: {
  dataRoot: string;
  seal: () => Promise<void>;
  restore: () => Promise<void>;
}): Promise<void> {
  initializeLocalDataProtection({ dataRoot: options.dataRoot });
  const result = await ensureLocalDataEncrypted({
    dataRoot: options.dataRoot,
    seal: options.seal,
    restore: options.restore,
  });
  if (result.migrated) {
    logger.info("Initial local data encryption migration completed", {
      contentKeyId: result.contentKeyId,
      finalizeWarning: result.finalizeWarning,
    });
  }
  if (result.finalizeWarning) {
    logger.warn("Encryption backup cleanup incomplete", {
      warning: result.finalizeWarning,
    });
  }
}

export async function afterDataPathMigrated(options: {
  previousDataPath: string;
  nextDataPath: string;
}): Promise<{ oldRootCleared: boolean; oldRootError?: string }> {
  registerManagedDataRoot(options.nextDataPath, "active_root");
  registerManagedDataRoot(options.previousDataPath, "old_root");
  return clearManagedDataRoot(options.previousDataPath);
}

export function afterUserDataExported(exportPath: string): void {
  registerManagedDataRoot(exportPath, "privacy_export");
}

export { rotateLocalDataKey };
