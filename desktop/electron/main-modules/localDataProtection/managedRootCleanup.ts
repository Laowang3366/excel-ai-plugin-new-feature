import * as fs from "node:fs";
import * as path from "node:path";
import { eraseManagedUserData } from "../userDataErase";
import {
  getLocalDataProtectionRuntime,
  registerManagedDataRoot,
} from "./localDataProtectionService";

/**
 * Clear a migrated old data root: managed dirs + settings + protection/export markers.
 * Failures leave the path registered as pending_erase.
 */
export async function clearManagedDataRoot(
  absolutePath: string,
): Promise<{ oldRootCleared: boolean; oldRootError?: string }> {
  const runtime = getLocalDataProtectionRuntime();
  try {
    const report = await eraseManagedUserData(absolutePath);
    if (report.errors.length > 0) {
      runtime?.registry.markStatus(absolutePath, "pending_erase", report.errors.join("; "));
      return { oldRootCleared: false, oldRootError: report.errors.join("; ") };
    }
    const settingsDir = path.join(absolutePath, "settings");
    if (fs.existsSync(settingsDir)) {
      const info = fs.lstatSync(settingsDir);
      if (info.isSymbolicLink()) throw new Error("拒绝删除符号链接或联接");
      fs.rmSync(settingsDir, { recursive: true, force: false });
    }
    for (const relative of [
      "privacy-export-manifest.json",
      path.join("sessions", ".local-data-protection.json"),
    ]) {
      const target = path.join(absolutePath, relative);
      if (!fs.existsSync(target)) continue;
      const info = fs.lstatSync(target);
      if (info.isSymbolicLink()) throw new Error("拒绝删除符号链接或联接");
      fs.rmSync(target, { force: false });
    }
    runtime?.registry.markStatus(absolutePath, "erased");
    runtime?.registry.removeErasedEntries();
    return { oldRootCleared: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    registerManagedDataRoot(absolutePath, "old_root");
    runtime?.registry.markStatus(absolutePath, "pending_erase", message);
    return { oldRootCleared: false, oldRootError: message };
  }
}
