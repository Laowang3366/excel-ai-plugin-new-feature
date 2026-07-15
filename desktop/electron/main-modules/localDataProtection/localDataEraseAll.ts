import * as fs from "node:fs";
import * as path from "node:path";
import { app } from "electron";
import { isPathInside, normalizePathForCompare } from "../settingsDataPath";
import { eraseManagedUserData, type UserDataEraseReport } from "../userDataErase";
import {
  buildEraseProofSummary,
  writeLocalDataEraseProof,
  type LocalDataEraseProof,
} from "./eraseProof";
import { createPayloadProtection, setPayloadProtection } from "./payloadProtection";
import {
  getLocalDataProtectionRuntime,
  type LocalDataProtectionRuntime,
} from "./localDataProtectionService";
import type { ManagedReplicaEntry } from "./managedReplicaRegistry";

const OLD_ROOT_EXTRA = ["settings"] as const;
const OLD_ROOT_FILES = [
  "privacy-export-manifest.json",
  "privacy-export-settings.json",
  path.join("sessions", ".local-data-protection.json"),
] as const;

export interface EraseAllManagedReplicasResult {
  success: boolean;
  erasedCategories: string[];
  errors: string[];
  keyMaterialDestroyed: boolean;
  proofSummary?: ReturnType<typeof buildEraseProofSummary>;
  proofPath?: string;
}

export async function eraseAllManagedReplicasAndKeys(
  options: {
    userDataPath?: string;
  } = {},
): Promise<EraseAllManagedReplicasResult> {
  const runtime = getLocalDataProtectionRuntime();
  if (!runtime) {
    return {
      success: false,
      erasedCategories: [],
      errors: ["本地数据保护未初始化"],
      keyMaterialDestroyed: false,
    };
  }
  const userDataPath = options.userDataPath ?? app.getPath("userData");
  const installIdDigest = runtime.keystore.installIdDigest();
  const replicas = runtime.registry.listErasable();
  const proofReplicas: LocalDataEraseProof["replicas"] = [];
  const erasedCategories: string[] = [];
  const errors: string[] = [];

  try {
    await assertErasableEntriesSafe(replicas);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return writeFailProof({
      installIdDigest,
      userDataPath,
      errors: [message],
      replicas: replicas.map((replica) => ({
        pathDigest: replica.pathDigest,
        category: replica.category,
        status: "skipped" as const,
        error: message,
      })),
      keyError: message,
    });
  }

  // Stage replacement key before any delete. Old active protection stays live.
  // Prior keys are everything except the replacement (handles leftover pending from a prior attempt).
  let replacementKeyId: number;
  let priorKeyIds: number[];
  try {
    replacementKeyId = runtime.keystore.createPendingKey();
    priorKeyIds = runtime.keystore.listKeyIds().filter((id) => id !== replacementKeyId);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return writeFailProof({
      installIdDigest,
      userDataPath,
      errors: [`replacement_key: ${message}`],
      replicas: replicas.map((replica) => ({
        pathDigest: replica.pathDigest,
        category: replica.category,
        status: "skipped" as const,
        error: "replacement_key_unavailable",
      })),
      keyError: message,
    });
  }

  for (const replica of replicas) {
    try {
      if (!fs.existsSync(replica.absolutePath)) {
        runtime.registry.markStatus(replica.absolutePath, "erased");
        proofReplicas.push({
          pathDigest: replica.pathDigest,
          category: replica.category,
          status: "erased",
        });
        erasedCategories.push(`${replica.category}:missing`);
        continue;
      }
      const report = await eraseReplicaByCategory(replica);
      if (report.errors.length > 0) {
        errors.push(...report.errors.map((item) => `${replica.pathDigest.slice(0, 8)}: ${item}`));
        runtime.registry.markStatus(
          replica.absolutePath,
          "pending_erase",
          report.errors.join("; "),
        );
        proofReplicas.push({
          pathDigest: replica.pathDigest,
          category: replica.category,
          status: "failed",
          error: report.errors.join("; "),
        });
        continue;
      }
      erasedCategories.push(
        ...report.erasedCategories.map((item) => `${replica.category}:${item}`),
      );
      runtime.registry.markStatus(replica.absolutePath, "erased");
      proofReplicas.push({
        pathDigest: replica.pathDigest,
        category: replica.category,
        status: "erased",
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      errors.push(`${replica.pathDigest.slice(0, 8)}: ${message}`);
      runtime.registry.markStatus(replica.absolutePath, "pending_erase", message);
      proofReplicas.push({
        pathDigest: replica.pathDigest,
        category: replica.category,
        status: "failed",
        error: message,
      });
    }
  }

  // Atomic promote replacement + purge all prior keys (single keystore write).
  let keyDestruction: LocalDataEraseProof["keyDestruction"] = {
    destroyedKeyIds: [],
    keyMaterialDestroyed: false,
  };
  try {
    const swapped = runtime.keystore.commitPendingKeyAndPurgePriorKeys();
    if (swapped.replacementKeyId !== replacementKeyId) {
      throw new Error("replacement_key_mismatch");
    }
    activateProtection(runtime);
    const remaining = runtime.keystore.listKeyIds();
    if (
      remaining.length !== 1 ||
      remaining[0] !== replacementKeyId ||
      !priorKeyIds.every((id) => swapped.destroyedKeyIds.includes(id))
    ) {
      throw new Error("prior_keys_not_fully_purged");
    }
    keyDestruction = {
      destroyedKeyIds: swapped.destroyedKeyIds,
      keyMaterialDestroyed: true,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    errors.push(`key_destruction: ${message}`);
    keyDestruction = {
      destroyedKeyIds: [],
      keyMaterialDestroyed: false,
      error: message,
    };
    try {
      activateProtection(runtime);
    } catch {
      /* keep previous global protection if refresh fails */
    }
    try {
      runtime.keystore.discardPendingKey();
    } catch {
      /* best-effort cleanup of unused pending */
    }
  }

  runtime.registry.removeErasedEntries();

  const proof: LocalDataEraseProof = {
    formatVersion: 1,
    createdAt: new Date().toISOString(),
    installIdDigest,
    keyDestruction,
    replicas: proofReplicas,
  };
  const written = writeLocalDataEraseProof(proof, userDataPath);
  return {
    success: errors.length === 0 && keyDestruction.keyMaterialDestroyed,
    erasedCategories,
    errors,
    keyMaterialDestroyed: keyDestruction.keyMaterialDestroyed,
    proofSummary: buildEraseProofSummary(proof, written.proofDigest),
    proofPath: written.proofPath,
  };
}

function activateProtection(runtime: LocalDataProtectionRuntime): void {
  const protection = createPayloadProtection(runtime.keystore);
  setPayloadProtection(protection);
  (runtime as { protection: typeof protection }).protection = protection;
}

function writeFailProof(input: {
  installIdDigest: string;
  userDataPath: string;
  errors: string[];
  replicas: LocalDataEraseProof["replicas"];
  keyError: string;
}): EraseAllManagedReplicasResult {
  const proof: LocalDataEraseProof = {
    formatVersion: 1,
    createdAt: new Date().toISOString(),
    installIdDigest: input.installIdDigest,
    keyDestruction: {
      destroyedKeyIds: [],
      keyMaterialDestroyed: false,
      error: input.keyError,
    },
    replicas: input.replicas,
  };
  const written = writeLocalDataEraseProof(proof, input.userDataPath);
  return {
    success: false,
    erasedCategories: [],
    errors: input.errors,
    keyMaterialDestroyed: false,
    proofSummary: buildEraseProofSummary(proof, written.proofDigest),
    proofPath: written.proofPath,
  };
}

async function eraseReplicaByCategory(replica: ManagedReplicaEntry): Promise<UserDataEraseReport> {
  if (replica.category === "active_root" || replica.category === "old_root") {
    const report = await eraseManagedUserData(replica.absolutePath);
    if (replica.category === "old_root") {
      const extra = await eraseOldRootExtras(replica.absolutePath);
      return {
        erasedCategories: [...report.erasedCategories, ...extra.erasedCategories],
        errors: [...report.errors, ...extra.errors],
      };
    }
    return report;
  }
  return erasePrivacyExportDirectory(replica.absolutePath);
}

async function eraseOldRootExtras(root: string): Promise<UserDataEraseReport> {
  const report: UserDataEraseReport = { erasedCategories: [], errors: [] };
  const resolved = path.resolve(root);
  for (const category of OLD_ROOT_EXTRA) {
    const target = path.join(resolved, category);
    try {
      const info = await fs.promises.lstat(target).catch((error: NodeJS.ErrnoException) => {
        if (error.code === "ENOENT") return null;
        throw error;
      });
      if (info?.isSymbolicLink()) throw new Error("拒绝删除符号链接或联接");
      if (info) await fs.promises.rm(target, { recursive: true, force: false });
      report.erasedCategories.push(category);
    } catch (error) {
      report.errors.push(`${category}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  for (const relative of OLD_ROOT_FILES) {
    const target = path.join(resolved, relative);
    try {
      const info = await fs.promises.lstat(target).catch((error: NodeJS.ErrnoException) => {
        if (error.code === "ENOENT") return null;
        throw error;
      });
      if (info?.isSymbolicLink()) throw new Error("拒绝删除符号链接或联接");
      if (info) await fs.promises.rm(target, { force: false });
      report.erasedCategories.push(relative);
    } catch (error) {
      report.errors.push(`${relative}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  return report;
}

async function erasePrivacyExportDirectory(exportPath: string): Promise<UserDataEraseReport> {
  const root = path.resolve(exportPath);
  if (path.dirname(root) === root) throw new Error("拒绝擦除磁盘根目录");
  const info = await fs.promises.lstat(root);
  if (!info.isDirectory() || info.isSymbolicLink()) {
    throw new Error("导出目录必须是非链接目录");
  }
  await fs.promises.rm(root, { recursive: true, force: false });
  return { erasedCategories: ["privacy_export_root"], errors: [] };
}

async function assertErasableEntriesSafe(replicas: ManagedReplicaEntry[]): Promise<void> {
  for (const replica of replicas) {
    const root = path.resolve(replica.absolutePath);
    if (path.dirname(root) === root) {
      throw new Error("拒绝擦除磁盘根目录");
    }
    if (fs.existsSync(root)) {
      const info = await fs.promises.lstat(root);
      if (info.isSymbolicLink()) {
        throw new Error(`拒绝擦除符号链接或联接: ${replica.pathDigest.slice(0, 8)}`);
      }
    }
  }
  for (let i = 0; i < replicas.length; i++) {
    for (let j = i + 1; j < replicas.length; j++) {
      const a = path.resolve(replicas[i]!.absolutePath);
      const b = path.resolve(replicas[j]!.absolutePath);
      if (normalizePathForCompare(a) === normalizePathForCompare(b)) continue;
      if (isPathInside(a, b) || isPathInside(b, a)) {
        throw new Error("拒绝擦除存在根包含关系的已登记路径");
      }
    }
  }
}
