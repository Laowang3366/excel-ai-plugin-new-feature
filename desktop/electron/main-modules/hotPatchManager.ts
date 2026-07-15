import { createHash, randomUUID } from "node:crypto";
import * as fs from "node:fs";
import { promises as fsp } from "node:fs";
import * as path from "node:path";

import type { HotPatchPolicy, HotPatchUpdate } from "./updateManifest";
import {
  extractHotPatchArchiveStreaming,
  isPathInside,
  MAX_ARCHIVE_BYTES,
  MAX_COMPRESSION_RATIO,
  MAX_ENTRY_BYTES,
  MAX_UNCOMPRESSED_BYTES,
  sha256File,
} from "./hotPatchArchive";

export { isAllowedHotPatchPath, sha256File } from "./hotPatchArchive";

const PATCH_STATE_FILE = "hot-patch-state.json";
const PATCH_SECURITY_STATE_FILE = "hot-patch-security-state.json";

interface HotPatchState {
  id: string;
  baseVersion: string;
  sequence: number;
  publishedAt: string;
  expiresAt: string;
  installedAt: string;
  rootPath: string;
  files: Array<{ path: string; sha256: string; size: number }>;
  healthStatus?: "pending" | "healthy";
  lastActivatedAt?: string;
}

interface HotPatchSecurityState {
  highestSequenceByBaseVersion: Record<string, number>;
  minimumSafeSequenceByBaseVersion: Record<string, number>;
  revokedPatchIds: string[];
}

function updatesRoot(userDataPath: string): string {
  return path.join(userDataPath, "updates");
}

function statePath(userDataPath: string): string {
  return path.join(updatesRoot(userDataPath), PATCH_STATE_FILE);
}

function patchRoot(userDataPath: string): string {
  return path.join(updatesRoot(userDataPath), "hot-patches");
}

function securityStatePath(userDataPath: string): string {
  return path.join(updatesRoot(userDataPath), PATCH_SECURITY_STATE_FILE);
}

async function writeStateAtomically(userDataPath: string, state: HotPatchState): Promise<void> {
  const target = statePath(userDataPath);
  await fsp.mkdir(path.dirname(target), { recursive: true });
  const temporary = `${target}.${randomUUID()}.tmp`;
  await fsp.writeFile(temporary, JSON.stringify(state, null, 2), "utf8");
  await fsp.rename(temporary, target);
}

function readSecurityState(userDataPath: string): HotPatchSecurityState {
  try {
    return JSON.parse(
      fs.readFileSync(securityStatePath(userDataPath), "utf8"),
    ) as HotPatchSecurityState;
  } catch {
    return {
      highestSequenceByBaseVersion: {},
      minimumSafeSequenceByBaseVersion: {},
      revokedPatchIds: [],
    };
  }
}

async function recordPatchSequence(
  userDataPath: string,
  baseVersion: string,
  sequence: number,
): Promise<void> {
  const state = readSecurityState(userDataPath);
  state.minimumSafeSequenceByBaseVersion ||= {};
  state.revokedPatchIds ||= [];
  state.highestSequenceByBaseVersion[baseVersion] = Math.max(
    state.highestSequenceByBaseVersion[baseVersion] ?? 0,
    sequence,
  );
  const target = securityStatePath(userDataPath);
  await fsp.mkdir(path.dirname(target), { recursive: true });
  const temporary = `${target}.${randomUUID()}.tmp`;
  await fsp.writeFile(temporary, JSON.stringify(state, null, 2), "utf8");
  await fsp.rename(temporary, target);
}

async function writeSecurityStateAtomically(
  userDataPath: string,
  state: HotPatchSecurityState,
): Promise<void> {
  const target = securityStatePath(userDataPath);
  await fsp.mkdir(path.dirname(target), { recursive: true });
  const temporary = `${target}.${randomUUID()}.tmp`;
  await fsp.writeFile(temporary, JSON.stringify(state, null, 2), "utf8");
  await fsp.rename(temporary, target);
}

function sha256FileSync(filePath: string): string {
  const hash = createHash("sha256");
  const descriptor = fs.openSync(filePath, "r");
  const buffer = Buffer.allocUnsafe(64 * 1024);
  try {
    let bytesRead = 0;
    while ((bytesRead = fs.readSync(descriptor, buffer, 0, buffer.length, null)) > 0) {
      hash.update(buffer.subarray(0, bytesRead));
    }
  } finally {
    fs.closeSync(descriptor);
  }
  return hash.digest("hex");
}

export async function installHotPatchArchive(input: {
  archivePath: string;
  descriptor: HotPatchUpdate;
  currentVersion: string;
  userDataPath: string;
}): Promise<HotPatchState> {
  const { archivePath, descriptor, currentVersion, userDataPath } = input;
  if (descriptor.baseVersion !== currentVersion) {
    throw new Error(`热补丁要求基础版本 ${descriptor.baseVersion}，当前版本为 ${currentVersion}`);
  }
  const now = Date.now();
  if (Date.parse(descriptor.publishedAt) > now + 5 * 60 * 1000)
    throw new Error("热补丁发布时间无效");
  if (Date.parse(descriptor.expiresAt) <= now) throw new Error("热补丁已过期");
  const highestSequence =
    readSecurityState(userDataPath).highestSequenceByBaseVersion[currentVersion] ?? 0;
  const securityState = readSecurityState(userDataPath);
  const minimumSafeSequence = securityState.minimumSafeSequenceByBaseVersion?.[currentVersion] ?? 0;
  if (securityState.revokedPatchIds?.includes(descriptor.id)) throw new Error("热补丁已被远程吊销");
  if (descriptor.sequence < minimumSafeSequence) throw new Error("热补丁序列低于远程安全基线");
  if (descriptor.sequence <= highestSequence) throw new Error("热补丁序列已安装或低于安全基线");
  const archiveStat = await fsp.stat(archivePath);
  if (archiveStat.size !== descriptor.size || archiveStat.size > MAX_ARCHIVE_BYTES) {
    throw new Error("热补丁归档大小无效");
  }
  const declaredTotal = descriptor.files.reduce((total, file) => total + file.size, 0);
  if (
    declaredTotal > MAX_UNCOMPRESSED_BYTES ||
    descriptor.files.some((file) => file.size > MAX_ENTRY_BYTES)
  ) {
    throw new Error("热补丁声明的解压体积超过限制");
  }
  if (declaredTotal / Math.max(1, archiveStat.size) > MAX_COMPRESSION_RATIO) {
    throw new Error("热补丁压缩比超过限制");
  }
  const actualHash = await sha256File(archivePath);
  if (actualHash.toLowerCase() !== descriptor.sha256.toLowerCase()) {
    throw new Error("热补丁文件哈希校验失败");
  }

  const expectedFiles = new Map(
    descriptor.files.map((file) => [file.path.replace(/\\/gu, "/"), file]),
  );
  if (expectedFiles.size !== descriptor.files.length) throw new Error("热补丁文件清单包含重复路径");

  const root = patchRoot(userDataPath);
  const staging = path.join(root, `.staging-${randomUUID()}`);
  const target = path.join(root, descriptor.id);
  await fsp.mkdir(staging, { recursive: true });

  try {
    await extractHotPatchArchiveStreaming(archivePath, staging, expectedFiles);

    await fsp.rm(target, { recursive: true, force: true });
    await fsp.rename(staging, target);
  } catch (error) {
    await fsp.rm(staging, { recursive: true, force: true });
    throw error;
  }

  const state: HotPatchState = {
    id: descriptor.id,
    baseVersion: descriptor.baseVersion,
    sequence: descriptor.sequence,
    publishedAt: descriptor.publishedAt,
    expiresAt: descriptor.expiresAt,
    installedAt: new Date().toISOString(),
    rootPath: target,
    files: descriptor.files,
    healthStatus: "pending",
  };
  await recordPatchSequence(userDataPath, descriptor.baseVersion, descriptor.sequence);
  await writeStateAtomically(userDataPath, state);
  return state;
}

export function activateInstalledHotPatch(
  currentVersion: string,
  userDataPath: string,
): string | null {
  try {
    const raw = fs.readFileSync(statePath(userDataPath), "utf8");
    const state = JSON.parse(raw) as HotPatchState;
    const expectedRoot = patchRoot(userDataPath);
    const securityState = readSecurityState(userDataPath);
    const minimumSafeSequence =
      securityState.minimumSafeSequenceByBaseVersion?.[currentVersion] ?? 0;
    if (
      state.baseVersion !== currentVersion ||
      securityState.revokedPatchIds?.includes(state.id) ||
      state.sequence < minimumSafeSequence ||
      (state.healthStatus === "pending" && Boolean(state.lastActivatedAt)) ||
      Date.parse(state.expiresAt) <= Date.now() ||
      !isPathInside(expectedRoot, path.resolve(state.rootPath)) ||
      !fs.existsSync(state.rootPath) ||
      !Array.isArray(state.files) ||
      state.files.some((file) => {
        const candidate = path.resolve(state.rootPath, file.path);
        return (
          !isPathInside(state.rootPath, candidate) ||
          !fs.existsSync(candidate) ||
          fs.statSync(candidate).size !== file.size ||
          sha256FileSync(candidate) !== file.sha256.toLowerCase()
        );
      })
    ) {
      delete process.env.WENGE_HOT_PATCH_ROOT;
      fs.rmSync(statePath(userDataPath), { force: true });
      return null;
    }
    state.healthStatus = "pending";
    state.lastActivatedAt = new Date().toISOString();
    writeStateAtomicallySync(userDataPath, state);
    process.env.WENGE_HOT_PATCH_ROOT = state.rootPath;
    return state.id;
  } catch {
    return null;
  }
}

function writeStateAtomicallySync(userDataPath: string, state: HotPatchState): void {
  const target = statePath(userDataPath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const temporary = `${target}.${randomUUID()}.tmp`;
  fs.writeFileSync(temporary, JSON.stringify(state, null, 2), "utf8");
  fs.renameSync(temporary, target);
}

export async function acknowledgeActiveHotPatchHealth(userDataPath: string): Promise<boolean> {
  try {
    const state = JSON.parse(await fsp.readFile(statePath(userDataPath), "utf8")) as HotPatchState;
    if (
      !process.env.WENGE_HOT_PATCH_ROOT ||
      path.resolve(process.env.WENGE_HOT_PATCH_ROOT) !== path.resolve(state.rootPath)
    ) {
      return false;
    }
    state.healthStatus = "healthy";
    delete state.lastActivatedAt;
    await writeStateAtomically(userDataPath, state);
    return true;
  } catch {
    return false;
  }
}

export async function applyHotPatchPolicy(
  userDataPath: string,
  currentVersion: string,
  policy: HotPatchPolicy | undefined,
): Promise<boolean> {
  if (!policy) return false;
  const state = readSecurityState(userDataPath);
  state.minimumSafeSequenceByBaseVersion = {
    ...(state.minimumSafeSequenceByBaseVersion || {}),
    ...policy.minimumSafeSequenceByBaseVersion,
  };
  state.revokedPatchIds = Array.from(
    new Set([...(state.revokedPatchIds || []), ...policy.revokedPatchIds]),
  ).sort();
  await writeSecurityStateAtomically(userDataPath, state);

  try {
    const active = JSON.parse(await fsp.readFile(statePath(userDataPath), "utf8")) as HotPatchState;
    const minimum = state.minimumSafeSequenceByBaseVersion[currentVersion] ?? 0;
    if (
      active.baseVersion === currentVersion &&
      (state.revokedPatchIds.includes(active.id) || active.sequence < minimum)
    ) {
      await disableActiveHotPatch(userDataPath);
      return true;
    }
  } catch {
    // No active patch state.
  }
  return false;
}

export function getActiveHotPatchId(userDataPath: string, currentVersion?: string): string | null {
  try {
    const state = JSON.parse(fs.readFileSync(statePath(userDataPath), "utf8")) as HotPatchState;
    if (currentVersion && state.baseVersion !== currentVersion) return null;
    return state.id || null;
  } catch {
    return null;
  }
}

export function resolveHotPatchPath(relativePath: string): string | null {
  const root = process.env.WENGE_HOT_PATCH_ROOT;
  if (!root) return null;
  const candidate = path.resolve(root, relativePath);
  if (!isPathInside(path.resolve(root), candidate) || !fs.existsSync(candidate)) return null;
  return candidate;
}

export async function disableActiveHotPatch(userDataPath: string): Promise<void> {
  delete process.env.WENGE_HOT_PATCH_ROOT;
  await fsp.rm(statePath(userDataPath), { force: true });
}
