import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { createHash, randomUUID } from "crypto";

import type { KnowledgeEntry, KnowledgeFileType, KnowledgeSource } from "./types";
import type { EmbeddingProfile } from "./embeddingService";
import { TextChunker } from "./textChunker";

export interface KnowledgeWriteInput {
  title?: string;
  content: string;
  tags?: string[];
  sourceName?: string;
  metadata?: Record<string, unknown>;
}

export interface KnowledgeWriteResult {
  sourcePath: string;
  sourceName: string;
  entryCount: number;
  entryIds: string[];
  indexedAt: number;
}

export type KnowledgeUpdateOperation = "replace" | "append";

export interface KnowledgeUpdateSourceInput {
  sourcePath: string;
  operation: KnowledgeUpdateOperation;
  content: string;
  title?: string;
  tags?: string[];
  metadata?: Record<string, unknown>;
}

export interface KnowledgeDeleteSourceInput {
  sourcePath: string;
}

export interface KnowledgeWriterOptions {
  notesDir?: string;
  maxTokens?: number;
}

interface KnowledgeWriteStore {
  bulkInsert(entries: KnowledgeEntry[]): void;
  upsertSource(source: KnowledgeSource): void;
  deleteSource?(sourcePath: string): void;
  getDbPath?(): string;
}

interface KnowledgeWriteEmbedder {
  embedBatch(texts: string[]): Promise<number[][]>;
  getProfile?: () => EmbeddingProfile;
}

export class KnowledgeWriter {
  private readonly store: KnowledgeWriteStore;
  private readonly embedder: KnowledgeWriteEmbedder;
  private readonly chunker: TextChunker;
  private readonly notesDir?: string;

  constructor(
    store: KnowledgeWriteStore,
    embedder: KnowledgeWriteEmbedder,
    options?: KnowledgeWriterOptions
  ) {
    this.store = store;
    this.embedder = embedder;
    this.chunker = new TextChunker(options?.maxTokens ?? 512);
    this.notesDir = options?.notesDir;
  }

  async writeNote(input: KnowledgeWriteInput): Promise<KnowledgeWriteResult> {
    const content = input.content.trim();
    if (!content) {
      throw new Error("知识库写入内容不能为空");
    }

    const now = Date.now();
    const id = randomUUID();
    const title = normalizeTitle(input.title) || makeTitleFromContent(content);
    const tags = normalizeTags(input.tags);
    const noteContent = buildNoteContent(title, content, tags, now);
    const notesDir = this.resolveNotesDir();
    fs.mkdirSync(notesDir, { recursive: true });

    const sourceName = buildSourceName(input.sourceName || title, id, now);
    const sourcePath = path.join(notesDir, sourceName);
    fs.writeFileSync(sourcePath, noteContent, "utf8");

    const result = await this.indexTextSource({
      sourcePath,
      sourceName,
      sourceType: "md",
      text: noteContent,
      metadata: {
        title,
        tags,
        createdAt: new Date(now).toISOString(),
        origin: "model",
        ...(input.metadata || {}),
      },
      indexedAt: now,
    });

    return {
      sourcePath,
      sourceName,
      entryCount: result.entryCount,
      entryIds: result.entryIds,
      indexedAt: result.indexedAt,
    };
  }

  async updateSource(input: KnowledgeUpdateSourceInput): Promise<KnowledgeWriteResult> {
    const sourcePath = normalizeSourcePath(input.sourcePath);
    const content = input.content.trim();
    if (!content) {
      throw new Error("knowledge update content cannot be empty");
    }
    if (input.operation !== "replace" && input.operation !== "append") {
      throw new Error("knowledge update operation must be replace or append");
    }

    const stat = fs.existsSync(sourcePath) ? fs.statSync(sourcePath) : null;
    if (!stat?.isFile()) {
      throw new Error(`knowledge source file does not exist: ${sourcePath}`);
    }

    const sourceType = getEditableSourceType(sourcePath);
    const sourceName = path.basename(sourcePath);
    const now = Date.now();
    const title = normalizeTitle(input.title) || path.parse(sourceName).name;
    const tags = normalizeTags(input.tags);
    const nextText =
      input.operation === "replace"
        ? buildEditableText(sourceType, title, content, tags, now)
        : appendEditableText(fs.readFileSync(sourcePath, "utf8"), content);

    fs.writeFileSync(sourcePath, nextText, "utf8");

    const result = await this.indexTextSource({
      sourcePath,
      sourceName,
      sourceType,
      text: nextText,
      metadata: {
        title,
        tags,
        updatedAt: new Date(now).toISOString(),
        origin: "model",
        operation: input.operation,
        ...(input.metadata || {}),
      },
      indexedAt: now,
    });

    return {
      sourcePath,
      sourceName,
      entryCount: result.entryCount,
      entryIds: result.entryIds,
      indexedAt: result.indexedAt,
    };
  }

  async deleteSource(input: KnowledgeDeleteSourceInput): Promise<{ sourcePath: string }> {
    const sourcePath = normalizeSourcePath(input.sourcePath);
    if (!sourcePath) {
      throw new Error("sourcePath cannot be empty");
    }
    if (!this.store.deleteSource) {
      throw new Error("knowledge store does not support deleting sources");
    }
    this.store.deleteSource(sourcePath);
    return { sourcePath };
  }

  private async indexTextSource(input: {
    sourcePath: string;
    sourceName: string;
    sourceType: Extract<KnowledgeFileType, "md" | "txt">;
    text: string;
    metadata: Record<string, unknown>;
    indexedAt: number;
  }): Promise<{ entryCount: number; entryIds: string[]; indexedAt: number }> {
    const chunks = this.chunker.chunk([
      {
        content: input.text,
        sourcePath: input.sourcePath,
        sourceName: input.sourceName,
        sourceType: input.sourceType,
        metadata: input.metadata,
      },
    ]);

    const embeddingResult = await this.embedChunks(chunks.map((chunk) => chunk.content));
    const embeddingProfile = embeddingResult.profile;
    const entries: KnowledgeEntry[] = chunks.map((chunk, index) => ({
      id: randomUUID(),
      source: "note",
      sourcePath: input.sourcePath,
      sourceName: input.sourceName,
      sourceType: input.sourceType,
      chunkIndex: chunk.index,
      content: chunk.content,
      metadata: chunk.metadata,
      embedding: embeddingResult.embeddings[index],
      embeddingProvider: embeddingResult.embeddings[index] ? embeddingProfile?.provider : undefined,
      embeddingModel: embeddingResult.embeddings[index] ? embeddingProfile?.model : undefined,
      embeddingDimensions: embeddingResult.embeddings[index]?.length ?? embeddingProfile?.dimensions,
      indexedAt: input.indexedAt,
      tokenCount: chunk.tokenCount,
    }));

    this.store.deleteSource?.(input.sourcePath);
    this.store.bulkInsert(entries);
    this.store.upsertSource({
      sourcePath: input.sourcePath,
      sourceName: input.sourceName,
      sourceType: input.sourceType,
      entryCount: entries.length,
      firstIndexed: input.indexedAt,
      lastIndexed: input.indexedAt,
      fileHash: createHash("sha256").update(input.text, "utf8").digest("hex").slice(0, 16),
    });

    return {
      entryCount: entries.length,
      entryIds: entries.map((entry) => entry.id),
      indexedAt: input.indexedAt,
    };
  }

  private resolveNotesDir(): string {
    if (this.notesDir) return this.notesDir;
    const dbPath = this.store.getDbPath?.();
    if (dbPath && dbPath !== ":memory:") {
      return path.join(path.dirname(dbPath), "notes");
    }
    return path.join(os.tmpdir(), "excel-ai-assistant-knowledge-notes");
  }

  private async embedChunks(texts: string[]): Promise<{
    embeddings: Array<number[] | null>;
    profile?: EmbeddingProfile;
  }> {
    try {
      const embeddings = await this.embedder.embedBatch(texts);
      return {
        embeddings,
        profile: this.embedder.getProfile?.(),
      };
    } catch {
      return {
        embeddings: texts.map(() => null),
      };
    }
  }
}

function buildNoteContent(title: string, content: string, tags: string[], timestamp: number): string {
  const lines = [
    `# ${title}`,
    "",
    `创建时间: ${new Date(timestamp).toISOString()}`,
  ];
  if (tags.length > 0) {
    lines.push(`标签: ${tags.join(", ")}`);
  }
  lines.push("", content);
  return lines.join("\n");
}

function buildEditableText(
  sourceType: Extract<KnowledgeFileType, "md" | "txt">,
  title: string,
  content: string,
  tags: string[],
  timestamp: number
): string {
  if (sourceType === "md") {
    return buildNoteContent(title, content, tags, timestamp);
  }
  return content;
}

function appendEditableText(existing: string, content: string): string {
  const trimmedExisting = existing.trimEnd();
  return `${trimmedExisting}${trimmedExisting ? "\n\n" : ""}${content}`;
}

function normalizeSourcePath(sourcePath: unknown): string {
  if (typeof sourcePath !== "string") return "";
  const trimmed = sourcePath.trim();
  return trimmed ? path.resolve(trimmed) : "";
}

function getEditableSourceType(sourcePath: string): Extract<KnowledgeFileType, "md" | "txt"> {
  const ext = path.extname(sourcePath).toLowerCase();
  if (ext === ".md") return "md";
  if (ext === ".txt") return "txt";
  throw new Error("only .md and .txt knowledge sources can be modified directly");
}

function normalizeTitle(title: unknown): string {
  return typeof title === "string" ? title.trim().slice(0, 120) : "";
}

function normalizeTags(tags: unknown): string[] {
  if (!Array.isArray(tags)) return [];
  const seen = new Set<string>();
  const result: string[] = [];
  for (const tag of tags) {
    if (typeof tag !== "string") continue;
    const normalized = tag.trim();
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized.slice(0, 40));
  }
  return result.slice(0, 20);
}

function makeTitleFromContent(content: string): string {
  const firstLine = content.split(/\r?\n/).map((line) => line.trim()).find(Boolean) || "知识条目";
  return firstLine.replace(/^#+\s*/, "").slice(0, 80) || "知识条目";
}

function buildSourceName(baseName: string, id: string, timestamp: number): string {
  const stamp = new Date(timestamp).toISOString().replace(/[-:T.Z]/g, "").slice(0, 14);
  const slug = sanitizeFileStem(baseName || "note").slice(0, 48) || "note";
  return `note-${stamp}-${id.slice(0, 8)}-${slug}.md`;
}

function sanitizeFileStem(value: string): string {
  return value
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
    .replace(/\s+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
}
