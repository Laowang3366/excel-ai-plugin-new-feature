import type { KnowledgeEntry, KnowledgeEntryRow, KnowledgeSource } from "./types";

export function entryToRow(entry: KnowledgeEntry): KnowledgeEntryRow {
  return {
    id: entry.id,
    source: entry.source,
    source_path: entry.sourcePath,
    source_name: entry.sourceName,
    source_type: entry.sourceType,
    chunk_index: entry.chunkIndex,
    content: entry.content,
    metadata: JSON.stringify(entry.metadata),
    embedding: entry.embedding ? JSON.stringify(entry.embedding) : null,
    embedding_provider: entry.embeddingProvider ?? null,
    embedding_model: entry.embeddingModel ?? null,
    embedding_dimensions: entry.embeddingDimensions ?? (entry.embedding ? entry.embedding.length : null),
    indexed_at: entry.indexedAt,
    token_count: entry.tokenCount,
  };
}

export function rowToEntry(row: Record<string, any>): KnowledgeEntry {
  return {
    id: row.id,
    source: row.source,
    sourcePath: row.source_path,
    sourceName: row.source_name,
    sourceType: row.source_type,
    chunkIndex: row.chunk_index,
    content: row.content,
    metadata: JSON.parse(row.metadata || "{}"),
    embedding: row.embedding ? JSON.parse(row.embedding) : null,
    embeddingProvider: row.embedding_provider ?? undefined,
    embeddingModel: row.embedding_model ?? undefined,
    embeddingDimensions: row.embedding_dimensions ?? undefined,
    indexedAt: row.indexed_at,
    tokenCount: row.token_count,
  };
}

export function rowToSource(row: Record<string, any>): KnowledgeSource {
  return {
    sourcePath: row.source_path,
    sourceName: row.source_name,
    sourceType: row.source_type,
    entryCount: row.entry_count,
    firstIndexed: row.first_indexed,
    lastIndexed: row.last_indexed,
    fileHash: row.file_hash,
  };
}

export function cosineSimilarity(a: Float64Array, b: Float64Array): number {
  if (a.length !== b.length) return 0;
  const len = a.length;
  let dotProduct = 0;
  let normA = 0;
  let normB = 0;
  for (let i = 0; i < len; i++) {
    dotProduct += a[i] * b[i];
    normA += a[i] * a[i];
    normB += b[i] * b[i];
  }
  const denom = Math.sqrt(normA) * Math.sqrt(normB);
  if (denom === 0) return 0;
  return dotProduct / denom;
}
