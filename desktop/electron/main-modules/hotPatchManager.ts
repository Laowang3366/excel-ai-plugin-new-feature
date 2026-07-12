import { createHash, randomUUID } from "node:crypto";
import * as fs from "node:fs";
import { promises as fsp } from "node:fs";
import * as path from "node:path";

import JSZip from "jszip";

import type { HotPatchUpdate } from "./updateManifest";

const PATCH_STATE_FILE = "hot-patch-state.json";
const MAX_PATCH_ENTRIES = 2_000;
const MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024;
const ALLOWED_PATCH_ROOTS = [
  "dist/",
  "public/knowledge/",
  "public/wps-jsa-bridge/",
];

interface HotPatchState {
  id: string;
  baseVersion: string;
  installedAt: string;
  rootPath: string;
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
  const actualHash = await sha256File(archivePath);
  if (actualHash.toLowerCase() !== descriptor.sha256.toLowerCase()) {
    throw new Error("热补丁文件哈希校验失败");
  }

  const zip = await JSZip.loadAsync(await fsp.readFile(archivePath));
  const entries = Object.values(zip.files).filter((entry) => !entry.dir);
  if (entries.length === 0 || entries.length > MAX_PATCH_ENTRIES) {
    throw new Error("热补丁文件数量无效");
  }
  for (const entry of entries) {
    if (!isAllowedHotPatchPath(entry.name)) {
      throw new Error(`热补丁包含不允许的文件: ${entry.name}`);
    }
  }
  validatePatchEntrySet(entries.map((entry) => entry.name));

  const root = patchRoot(userDataPath);
  const staging = path.join(root, `.staging-${randomUUID()}`);
  const target = path.join(root, descriptor.id);
  await fsp.mkdir(staging, { recursive: true });

  let totalBytes = 0;
  try {
    for (const entry of entries) {
      const normalized = entry.name.replace(/\\/gu, "/");
      const destination = path.resolve(staging, normalized);
      if (!isPathInside(staging, destination)) throw new Error(`热补丁路径越界: ${entry.name}`);
      const data = await entry.async("nodebuffer");
      totalBytes += data.byteLength;
      if (totalBytes > MAX_UNCOMPRESSED_BYTES) throw new Error("热补丁解压后体积超过限制");
      await fsp.mkdir(path.dirname(destination), { recursive: true });
      await fsp.writeFile(destination, data);
    }

    await fsp.rm(target, { recursive: true, force: true });
    await fsp.rename(staging, target);
  } catch (error) {
    await fsp.rm(staging, { recursive: true, force: true });
    throw error;
  }

  const state: HotPatchState = {
    id: descriptor.id,
    baseVersion: descriptor.baseVersion,
    installedAt: new Date().toISOString(),
    rootPath: target,
  };
  await writeStateAtomically(userDataPath, state);
  return state;
}

export function activateInstalledHotPatch(currentVersion: string, userDataPath: string): string | null {
  try {
    const raw = fs.readFileSync(statePath(userDataPath), "utf8");
    const state = JSON.parse(raw) as HotPatchState;
    const expectedRoot = patchRoot(userDataPath);
    if (
      state.baseVersion !== currentVersion ||
      !isPathInside(expectedRoot, path.resolve(state.rootPath)) ||
      !fs.existsSync(state.rootPath)
    ) {
      return null;
    }
    process.env.WENGE_HOT_PATCH_ROOT = state.rootPath;
    return state.id;
  } catch {
    return null;
  }
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
