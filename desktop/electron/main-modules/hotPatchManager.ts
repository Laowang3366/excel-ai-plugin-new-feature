import { createHash, randomUUID } from "node:crypto";
import * as fs from "node:fs";
import { promises as fsp } from "node:fs";
import * as path from "node:path";

import { Unzip, UnzipInflate, UnzipPassThrough, type UnzipFile } from "fflate";

import type { HotPatchPolicy, HotPatchUpdate } from "./updateManifest";

const PATCH_STATE_FILE = "hot-patch-state.json";
const PATCH_SECURITY_STATE_FILE = "hot-patch-security-state.json";
const MAX_PATCH_ENTRIES = 2_000;
const MAX_ARCHIVE_BYTES = 50 * 1024 * 1024;
const MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024;
const MAX_ENTRY_BYTES = 25 * 1024 * 1024;
const MAX_COMPRESSION_RATIO = 100;
const ALLOWED_PATCH_ROOTS = [
  "dist/",
  "public/knowledge/",
  "public/wps-jsa-bridge/",
];

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

function isPathInside(parent: string, candidate: string): boolean {
  const relative = path.relative(parent, candidate);
  return relative !== "" && !relative.startsWith("..") && !path.isAbsolute(relative);
}

export function isAllowedHotPatchPath(entryName: string): boolean {
  const normalized = entryName.replace(/\\/gu, "/").replace(/^\.\//u, "");
  if (!normalized || normalized.startsWith("/") || normalized.includes("../")) return false;
  return ALLOWED_PATCH_ROOTS.some((root) => normalized.startsWith(root));
}

function validatePatchEntrySet(entryNames: string[]): void {
  const entries = new Set(entryNames.map((name) => name.replace(/\\/gu, "/")));
  if ([...entries].some((name) => name.startsWith("dist/")) && !entries.has("dist/index.html")) {
    throw new Error("界面热补丁缺少 dist/index.html");
  }
  if (
    [...entries].some((name) => name.startsWith("public/knowledge/")) &&
    !entries.has("public/knowledge/builtin-knowledge.json")
  ) {
    throw new Error("知识库热补丁缺少 builtin-knowledge.json");
  }
  if (
    [...entries].some((name) => name.startsWith("public/wps-jsa-bridge/")) &&
    !entries.has("public/wps-jsa-bridge/index.html")
  ) {
    throw new Error("WPS 桥接热补丁缺少 index.html");
  }
}

export async function sha256File(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
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
    return JSON.parse(fs.readFileSync(securityStatePath(userDataPath), "utf8")) as HotPatchSecurityState;
  } catch {
    return {
      highestSequenceByBaseVersion: {},
      minimumSafeSequenceByBaseVersion: {},
      revokedPatchIds: [],
    };
  }
}

async function recordPatchSequence(userDataPath: string, baseVersion: string, sequence: number): Promise<void> {
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
  if (Date.parse(descriptor.publishedAt) > now + 5 * 60 * 1000) throw new Error("热补丁发布时间无效");
  if (Date.parse(descriptor.expiresAt) <= now) throw new Error("热补丁已过期");
  const highestSequence = readSecurityState(userDataPath).highestSequenceByBaseVersion[currentVersion] ?? 0;
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
  if (declaredTotal > MAX_UNCOMPRESSED_BYTES || descriptor.files.some((file) => file.size > MAX_ENTRY_BYTES)) {
    throw new Error("热补丁声明的解压体积超过限制");
  }
  if (declaredTotal / Math.max(1, archiveStat.size) > MAX_COMPRESSION_RATIO) {
    throw new Error("热补丁压缩比超过限制");
  }
  const actualHash = await sha256File(archivePath);
  if (actualHash.toLowerCase() !== descriptor.sha256.toLowerCase()) {
    throw new Error("热补丁文件哈希校验失败");
  }

  const expectedFiles = new Map(descriptor.files.map((file) => [file.path.replace(/\\/gu, "/"), file]));
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

async function extractHotPatchArchiveStreaming(
  archivePath: string,
  staging: string,
  expectedFiles: Map<string, { path: string; sha256: string; size: number }>,
): Promise<void> {
  let entryCount = 0;
  let completedCount = 0;
  let totalBytes = 0;
  let fatalError: Error | null = null;
  const seen = new Set<string>();
  const openDescriptors = new Set<number>();

  const fail = (message: string, file?: UnzipFile): void => {
    fatalError ||= new Error(message);
    file?.terminate();
  };

  const unzip = new Unzip((file) => {
    entryCount += 1;
    const normalized = file.name.replace(/\\/gu, "/");
    if (entryCount > MAX_PATCH_ENTRIES) return fail("热补丁文件数量无效", file);
    if (!isAllowedHotPatchPath(normalized)) return fail(`热补丁包含不允许的文件: ${file.name}`, file);
    if (seen.has(normalized)) return fail(`热补丁包含重复文件: ${file.name}`, file);
    const expected = expectedFiles.get(normalized);
    if (!expected) return fail(`热补丁文件清单与归档不一致: ${file.name}`, file);
    if (file.originalSize !== undefined && file.originalSize !== expected.size) {
      return fail(`热补丁文件声明大小不一致: ${file.name}`, file);
    }
    if (file.originalSize !== undefined && file.originalSize > MAX_ENTRY_BYTES) {
      return fail(`热补丁文件超过单文件限制: ${file.name}`, file);
    }
    if (
      file.size !== undefined &&
      file.originalSize !== undefined &&
      file.originalSize / Math.max(1, file.size) > MAX_COMPRESSION_RATIO
    ) {
      return fail(`热补丁文件压缩比超过限制: ${file.name}`, file);
    }

    const destination = path.resolve(staging, normalized);
    if (!isPathInside(staging, destination)) return fail(`热补丁路径越界: ${file.name}`, file);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    const descriptor = fs.openSync(destination, "wx");
    openDescriptors.add(descriptor);
    const hash = createHash("sha256");
    let entryBytes = 0;
    seen.add(normalized);

    file.ondata = (error, data, final) => {
      if (fatalError) return;
      if (error) return fail(`热补丁解压失败: ${file.name}: ${error.message}`, file);
      entryBytes += data.byteLength;
      totalBytes += data.byteLength;
      if (entryBytes > expected.size || entryBytes > MAX_ENTRY_BYTES) {
        return fail(`热补丁文件解压大小超过限制: ${file.name}`, file);
      }
      if (totalBytes > MAX_UNCOMPRESSED_BYTES) {
        return fail("热补丁解压后体积超过限制", file);
      }
      if (data.byteLength > 0) {
        fs.writeSync(descriptor, data);
        hash.update(data);
      }
      if (!final) return;
      fs.closeSync(descriptor);
      openDescriptors.delete(descriptor);
      if (entryBytes !== expected.size || hash.digest("hex") !== expected.sha256.toLowerCase()) {
        return fail(`热补丁文件校验失败: ${file.name}`, file);
      }
      completedCount += 1;
    };
    file.start();
  });
  unzip.register(UnzipPassThrough);
  unzip.register(UnzipInflate);

  try {
    const input = fs.createReadStream(archivePath, { highWaterMark: 64 * 1024 });
    for await (const chunk of input) {
      unzip.push(chunk as Buffer, false);
      if (fatalError) throw fatalError;
    }
    unzip.push(new Uint8Array(), true);
    if (fatalError) throw fatalError;
    if (
      entryCount === 0 ||
      entryCount !== expectedFiles.size ||
      completedCount !== expectedFiles.size ||
      seen.size !== expectedFiles.size
    ) {
      throw new Error("热补丁文件清单与归档不一致");
    }
    validatePatchEntrySet([...seen]);
  } finally {
    for (const descriptor of openDescriptors) {
      try {
        fs.closeSync(descriptor);
      } catch {
        /* descriptor already closed */
      }
    }
  }
}

export function activateInstalledHotPatch(currentVersion: string, userDataPath: string): string | null {
  try {
    const raw = fs.readFileSync(statePath(userDataPath), "utf8");
    const state = JSON.parse(raw) as HotPatchState;
    const expectedRoot = patchRoot(userDataPath);
    const securityState = readSecurityState(userDataPath);
    const minimumSafeSequence = securityState.minimumSafeSequenceByBaseVersion?.[currentVersion] ?? 0;
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
        return !isPathInside(state.rootPath, candidate) || !fs.existsSync(candidate) ||
          fs.statSync(candidate).size !== file.size || sha256FileSync(candidate) !== file.sha256.toLowerCase();
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
    if (!process.env.WENGE_HOT_PATCH_ROOT || path.resolve(process.env.WENGE_HOT_PATCH_ROOT) !== path.resolve(state.rootPath)) {
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
  state.revokedPatchIds = Array.from(new Set([
    ...(state.revokedPatchIds || []),
    ...policy.revokedPatchIds,
  ])).sort();
  await writeSecurityStateAtomically(userDataPath, state);

  try {
    const active = JSON.parse(await fsp.readFile(statePath(userDataPath), "utf8")) as HotPatchState;
    const minimum = state.minimumSafeSequenceByBaseVersion[currentVersion] ?? 0;
    if (active.baseVersion === currentVersion && (
      state.revokedPatchIds.includes(active.id) || active.sequence < minimum
    )) {
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
