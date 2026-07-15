import type { Dirent } from "node:fs";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as zlib from "node:zlib";
import { promisify } from "node:util";
import type { RolloutItem, RolloutLine } from "../shared/types";
import { clampNumber } from "../shared/numberLimits";
import { getRolloutTurnId } from "./stateRuntimeMappers";
import { extractRolloutSearchContent } from "./rolloutSearchContent";

const gzipAsync = promisify(zlib.gzip);

export interface RolloutArchiveOptions {
  archiveRolloutAfterBytes?: number;
}

export type RolloutArchiveResult =
  | { archived: true; archivePath: string; bytesBefore: number; bytesAfter: number }
  | { archived: false; reason: "disabled" | "below_threshold" | "missing" };

export interface RolloutCompressionWorkerOptions {
  sessionsRoot: string;
  activeThreadIds?: Iterable<string>;
  coldAfterMs?: number;
  minBytes?: number;
  now?: () => number;
}

export interface RolloutCompressionEntry {
  sourcePath: string;
  archivePath: string;
  bytesBefore: number;
  bytesAfter: number;
}

export interface RolloutCompressionSkipped {
  sourcePath: string;
  reason: "active_thread" | "not_cold" | "below_threshold" | "already_compressed" | "missing";
}

export interface RolloutCompressionWorkerResult {
  compressed: RolloutCompressionEntry[];
  skipped: RolloutCompressionSkipped[];
}

export interface CompressedRolloutSearchOptions {
  sessionsRoot: string;
  query: string;
  limit?: number;
}

export interface CompressedRolloutSearchMatch {
  id: number;
  threadId: string;
  turnId?: string;
  itemType: string;
  timestamp: string;
  item: RolloutItem;
  snippet: string;
  archivePath: string;
}

/**
 * 为过大的 rollout JSONL 创建 gzip 快照。
 *
 * 关联模块：
 * - sessionStore.ts: 管理活跃 JSONL，本模块不删除源文件，避免破坏恢复链路。
 * - core/agentLoop: 可在压缩完成后按配置触发归档快照。
 */
export async function archiveRolloutSnapshotIfNeeded(
  filePath: string,
  options: RolloutArchiveOptions,
): Promise<RolloutArchiveResult> {
  const threshold = options.archiveRolloutAfterBytes;
  if (!threshold || threshold <= 0) return { archived: false, reason: "disabled" };

  let stat;
  try {
    stat = await fs.stat(filePath);
  } catch {
    return { archived: false, reason: "missing" };
  }

  if (stat.size < threshold) return { archived: false, reason: "below_threshold" };

  const content = await fs.readFile(filePath);
  const compressed = await gzipAsync(content);
  const archivePath = `${filePath}.gz`;
  await fs.writeFile(archivePath, compressed);

  return {
    archived: true,
    archivePath,
    bytesBefore: stat.size,
    bytesAfter: compressed.byteLength,
  };
}

/**
 * 后台压缩冷 rollout JSONL。
 *
 * 关联模块：
 * - sessionStore.ts: 提供 sessionsRoot 与活跃 threadId；本模块不删除源 JSONL，保持恢复兼容。
 * - stateRuntimeStore.ts: logs.db 才是查询主存储，JSONL 作为审计/回放副本可异步压缩。
 */
export async function spawnRolloutCompressionWorker(
  options: RolloutCompressionWorkerOptions,
): Promise<RolloutCompressionWorkerResult> {
  const activeThreadIds = new Set(options.activeThreadIds ?? []);
  const coldAfterMs = options.coldAfterMs ?? 24 * 60 * 60 * 1000;
  const minBytes = options.minBytes ?? 1024 * 1024;
  const now = options.now?.() ?? Date.now();
  const files = await collectRolloutJsonlFiles(options.sessionsRoot);
  const compressed: RolloutCompressionEntry[] = [];
  const skipped: RolloutCompressionSkipped[] = [];

  for (const sourcePath of files) {
    const archivePath = `${sourcePath}.zst`;
    const threadId = getThreadIdFromRolloutPath(sourcePath);
    if (threadId && activeThreadIds.has(threadId)) {
      skipped.push({ sourcePath, reason: "active_thread" });
      continue;
    }

    let sourceStat;
    try {
      sourceStat = await fs.stat(sourcePath);
    } catch {
      skipped.push({ sourcePath, reason: "missing" });
      continue;
    }

    if (now - sourceStat.mtimeMs < coldAfterMs) {
      skipped.push({ sourcePath, reason: "not_cold" });
      continue;
    }
    if (sourceStat.size < minBytes) {
      skipped.push({ sourcePath, reason: "below_threshold" });
      continue;
    }
    if (await exists(archivePath)) {
      skipped.push({ sourcePath, reason: "already_compressed" });
      continue;
    }

    const content = await fs.readFile(sourcePath);
    const compressedContent = await zstdCompressAsync(content);
    await fs.writeFile(archivePath, compressedContent);
    compressed.push({
      sourcePath,
      archivePath,
      bytesBefore: sourceStat.size,
      bytesAfter: compressedContent.byteLength,
    });
  }

  return { compressed, skipped };
}

/** 搜索已 zstd 压缩的冷 rollout JSONL 归档。 */
export async function searchCompressedRolloutMatches(
  options: CompressedRolloutSearchOptions,
): Promise<CompressedRolloutSearchMatch[]> {
  const terms = normalizeQueryTerms(options.query);
  if (terms.length === 0) return [];

  const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
  const archives = await collectCompressedRolloutFiles(options.sessionsRoot);
  const matches: CompressedRolloutSearchMatch[] = [];
  let nextId = -1;

  for (const archivePath of archives) {
    const threadId = getThreadIdFromRolloutPath(archivePath);
    if (!threadId) continue;

    const content = await zstdDecompressAsync(await fs.readFile(archivePath));
    const relative = path
      .relative(options.sessionsRoot, archivePath.replace(/\.zst$/u, ""))
      .split(path.sep)
      .join("/");
    for (const line of content.toString("utf-8").split(/\r?\n/)) {
      if (!line.trim()) continue;
      let parsed: RolloutLine;
      try {
        const { getPayloadProtection } =
          await import("../../main-modules/localDataProtection/payloadProtection");
        const { isProtectedBlob, jsonlLineAad, parseProtectedRecordId } =
          await import("../../main-modules/localDataProtection/protectedBlob");
        let plain = line;
        if (isProtectedBlob(line)) {
          const rid = parseProtectedRecordId(line);
          const protection = getPayloadProtection();
          if (!rid || !protection) continue;
          plain = protection.unprotect(line, jsonlLineAad(relative, rid));
        }
        parsed = JSON.parse(plain) as RolloutLine;
      } catch {
        continue;
      }

      const searchable = extractRolloutSearchContent(parsed.item);
      if (!matchesQuery(searchable, terms)) continue;
      matches.push({
        id: nextId--,
        threadId,
        turnId: getRolloutTurnId(parsed.item) ?? undefined,
        itemType: parsed.item.type,
        timestamp: parsed.timestamp,
        item: parsed.item,
        snippet: buildSnippet(searchable, terms),
        archivePath,
      });
      if (matches.length >= limit) return matches;
    }
  }

  return matches;
}

async function collectRolloutJsonlFiles(root: string): Promise<string[]> {
  const files: string[] = [];
  await collectRolloutJsonlFilesInto(root, files);
  return files;
}

async function collectCompressedRolloutFiles(root: string): Promise<string[]> {
  const files: string[] = [];
  await collectCompressedRolloutFilesInto(root, files);
  return files;
}

async function collectCompressedRolloutFilesInto(dir: string, files: string[]): Promise<void> {
  let entries: Dirent[];
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectCompressedRolloutFilesInto(fullPath, files);
    } else if (
      entry.isFile() &&
      entry.name.startsWith("rollout-") &&
      entry.name.endsWith(".jsonl.zst")
    ) {
      files.push(fullPath);
    }
  }
}

async function collectRolloutJsonlFilesInto(dir: string, files: string[]): Promise<void> {
  let entries: Dirent[];
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectRolloutJsonlFilesInto(fullPath, files);
    } else if (
      entry.isFile() &&
      entry.name.startsWith("rollout-") &&
      entry.name.endsWith(".jsonl")
    ) {
      files.push(fullPath);
    }
  }
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function zstdCompressAsync(content: Buffer): Promise<Buffer> {
  const compress = (
    zlib as typeof zlib & {
      zstdCompress?: (
        buffer: Buffer,
        callback: (error: Error | null, result: Buffer) => void,
      ) => void;
    }
  ).zstdCompress;
  if (!compress) {
    throw new Error("当前 Node 运行时不支持 zstd 压缩");
  }

  return new Promise((resolve, reject) => {
    compress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}

async function zstdDecompressAsync(content: Buffer): Promise<Buffer> {
  const decompress = (
    zlib as typeof zlib & {
      zstdDecompress?: (
        buffer: Buffer,
        callback: (error: Error | null, result: Buffer) => void,
      ) => void;
    }
  ).zstdDecompress;
  if (!decompress) {
    throw new Error("当前 Node 运行时不支持 zstd 解压");
  }

  return new Promise((resolve, reject) => {
    decompress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}

function getThreadIdFromRolloutPath(filePath: string): string | null {
  const normalizedPath = filePath.endsWith(".zst") ? filePath.slice(0, -4) : filePath;
  const filename = path.basename(normalizedPath, ".jsonl");
  const match = filename.match(/(thread-[A-Za-z0-9_-]+)$/);
  return match?.[1] ?? null;
}

function normalizeQueryTerms(query: string): string[] {
  return query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
}

function matchesQuery(content: string, terms: string[]): boolean {
  const normalized = content.toLocaleLowerCase();
  return terms.every((term) => normalized.includes(term));
}

function buildSnippet(content: string, terms: string[]): string {
  const normalized = content.toLocaleLowerCase();
  const firstIndex =
    terms
      .map((term) => normalized.indexOf(term))
      .filter((index) => index >= 0)
      .sort((a, b) => a - b)[0] ?? 0;
  const start = Math.max(0, firstIndex - 24);
  const end = Math.min(content.length, firstIndex + 80);
  return `${start > 0 ? "..." : ""}${content.slice(start, end)}${end < content.length ? "..." : ""}`;
}
