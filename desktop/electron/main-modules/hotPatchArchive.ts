import { createHash } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";

import { Unzip, UnzipInflate, UnzipPassThrough, type UnzipFile } from "fflate";

export const MAX_PATCH_ENTRIES = 2_000;
export const MAX_ARCHIVE_BYTES = 50 * 1024 * 1024;
export const MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024;
export const MAX_ENTRY_BYTES = 25 * 1024 * 1024;
export const MAX_COMPRESSION_RATIO = 100;

const ALLOWED_PATCH_ROOTS = ["dist/", "public/knowledge/", "public/wps-jsa-bridge/"];

export function isPathInside(parent: string, candidate: string): boolean {
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

export async function extractHotPatchArchiveStreaming(
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
    if (!isAllowedHotPatchPath(normalized)) {
      return fail(`热补丁包含不允许的文件: ${file.name}`, file);
    }
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
        // Descriptor was already closed by the unzip callback.
      }
    }
  }
}
