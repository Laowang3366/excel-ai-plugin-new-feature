import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { createHash, randomUUID } from "crypto";

import type { KnowledgeEntry, KnowledgeSource } from "./types";
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

    const chunks = this.chunker.chunk([
      {
        content: noteContent,
        sourcePath,
        sourceName,
        sourceType: "md",
        metadata: {
          title,
          tags,
          createdAt: new Date(now).toISOString(),
          origin: "model",
          ...(input.metadata || {}),
        },
      },
    ]);

    const embeddings = await this.embedder.embedBatch(chunks.map((chunk) => chunk.content));
    const embeddingProfile = this.embedder.getProfile?.();
    const entries: KnowledgeEntry[] = chunks.map((chunk, index) => ({
      id: randomUUID(),
      source: "note",
      sourcePath,
      sourceName,
      sourceType: "md",
      chunkIndex: chunk.index,
      content: chunk.content,
      metadata: chunk.metadata,
      embedding: embeddings[index],
      embeddingProvider: embeddingProfile?.provider,
      embeddingModel: embeddingProfile?.model,
      embeddingDimensions: embeddingProfile?.dimensions ?? embeddings[index]?.length,
      indexedAt: now,
      tokenCount: chunk.tokenCount,
    }));

    this.store.deleteSource?.(sourcePath);
    this.store.bulkInsert(entries);
    this.store.upsertSource({
      sourcePath,
      sourceName,
      sourceType: "md",
      entryCount: entries.length,
      firstIndexed: now,
      lastIndexed: now,
      fileHash: createHash("sha256").update(noteContent, "utf8").digest("hex").slice(0, 16),
    });

    return {
      sourcePath,
      sourceName,
      entryCount: entries.length,
      entryIds: entries.map((entry) => entry.id),
      indexedAt: now,
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
