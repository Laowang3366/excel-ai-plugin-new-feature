import type { EmbeddingProfile } from "./embeddingService";
import type { KnowledgeEntry, KnowledgeResult } from "./types";
import { cosineSimilarity, rowToEntry } from "./sqliteStoreRows";
import type { SqliteDatabase } from "../storage/nodeSqlite";

type SearchFilter = {
  sourceFilter?: string[];
  pathFilter?: string[];
};

type VectorSearchFilter = SearchFilter & {
  embeddingProfile?: EmbeddingProfile;
};

function appendListFilter(
  sql: string,
  params: Array<string | number>,
  column: "source" | "source_path",
  values: string[] | undefined,
): string {
  if (!values || values.length === 0) return sql;

  params.push(...values);
  return `${sql} AND ${column} IN (${values.map(() => "?").join(",")})`;
}

export function searchKnowledgeByVector(
  db: SqliteDatabase,
  queryVector: number[],
  topK: number,
  filter?: VectorSearchFilter,
): KnowledgeResult[] {
  let sql = "SELECT * FROM knowledge_entries WHERE embedding IS NOT NULL";
  const params: Array<string | number> = [];

  if (filter?.embeddingProfile) {
    sql += " AND embedding_provider = ? AND embedding_model = ? AND embedding_dimensions = ?";
    params.push(
      filter.embeddingProfile.provider,
      filter.embeddingProfile.model,
      filter.embeddingProfile.dimensions,
    );
  }

  sql = appendListFilter(sql, params, "source", filter?.sourceFilter);
  sql = appendListFilter(sql, params, "source_path", filter?.pathFilter);

  const rows = db.prepare(sql).all(...params) as Record<string, any>[];
  const queryVec = new Float64Array(queryVector);
  const results: KnowledgeResult[] = [];

  for (const row of rows) {
    try {
      const entryVec = new Float64Array(JSON.parse(row.embedding));
      const score = cosineSimilarity(queryVec, entryVec);
      if (score > 0) results.push({ entry: rowToEntry(row), score });
    } catch {
      // A corrupt row must not prevent healthy knowledge entries from being searched.
    }
  }

  results.sort((a, b) => b.score - a.score);
  return results.slice(0, topK);
}

export function searchKnowledgeByKeyword(
  db: SqliteDatabase,
  keywords: string[],
  topK: number,
  filter?: SearchFilter,
): KnowledgeEntry[] {
  if (keywords.length === 0) return [];

  const seen = new Set<string>();
  const results: KnowledgeEntry[] = [];

  for (const keyword of keywords) {
    let sql = "SELECT * FROM knowledge_entries WHERE content LIKE ?";
    const params: Array<string | number> = [`%${keyword}%`];

    sql = appendListFilter(sql, params, "source", filter?.sourceFilter);
    sql = appendListFilter(sql, params, "source_path", filter?.pathFilter);
    params.push(topK);

    const rows = db.prepare(`${sql} LIMIT ?`).all(...params) as Record<string, any>[];
    for (const row of rows) {
      if (seen.has(row.id)) continue;

      seen.add(row.id);
      results.push(rowToEntry(row));
      if (results.length >= topK) break;
    }
    if (results.length >= topK) break;
  }

  return results.slice(0, topK);
}
