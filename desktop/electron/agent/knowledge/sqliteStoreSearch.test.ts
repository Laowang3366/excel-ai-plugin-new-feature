import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { openSqliteDatabase } from "../storage/nodeSqlite";
import type { SqliteDatabase } from "../storage/nodeSqlite";
import { initKnowledgeTables } from "./sqliteStoreSchema";
import { searchKnowledgeByKeyword, searchKnowledgeByVector } from "./sqliteStoreSearch";

describe("sqliteStoreSearch", () => {
  let db: SqliteDatabase;

  beforeEach(() => {
    db = openSqliteDatabase(":memory:");
    initKnowledgeTables(db);
  });

  afterEach(() => {
    db.close();
  });

  it("combines source and path filters for vector and keyword search", () => {
    insertEntry(db, {
      id: "matching",
      source: "document",
      sourcePath: "/knowledge/matching.md",
      content: "季度销售分析",
      embedding: "[1,0]",
    });
    insertEntry(db, {
      id: "wrong-path",
      source: "document",
      sourcePath: "/knowledge/other.md",
      content: "季度销售分析",
      embedding: "[1,0]",
    });
    insertEntry(db, {
      id: "wrong-source",
      source: "workbook",
      sourcePath: "/knowledge/matching.md",
      content: "季度销售分析",
      embedding: "[1,0]",
    });

    const filter = {
      sourceFilter: ["document"],
      pathFilter: ["/knowledge/matching.md"],
    };

    expect(
      searchKnowledgeByVector(db, [1, 0], 10, filter).map((result) => result.entry.id),
    ).toEqual(["matching"]);
    expect(searchKnowledgeByKeyword(db, ["销售"], 10, filter).map((entry) => entry.id)).toEqual([
      "matching",
    ]);
  });

  it("skips a corrupt embedding row without hiding healthy results", () => {
    insertEntry(db, {
      id: "corrupt",
      source: "document",
      sourcePath: "/knowledge/corrupt.md",
      content: "损坏向量",
      embedding: "not-json",
    });
    insertEntry(db, {
      id: "healthy",
      source: "document",
      sourcePath: "/knowledge/healthy.md",
      content: "健康向量",
      embedding: "[1,0]",
    });

    const results = searchKnowledgeByVector(db, [1, 0], 10);

    expect(results.map((result) => result.entry.id)).toEqual(["healthy"]);
  });
});

function insertEntry(
  db: SqliteDatabase,
  entry: {
    id: string;
    source: "document" | "workbook";
    sourcePath: string;
    content: string;
    embedding: string;
  },
): void {
  db.prepare(
    `INSERT INTO knowledge_entries
      (id, source, source_path, source_name, source_type, chunk_index,
       content, metadata, embedding, indexed_at, token_count)
     VALUES (?, ?, ?, ?, ?, 0, ?, '{}', ?, 1, 1)`,
  ).run(
    entry.id,
    entry.source,
    entry.sourcePath,
    entry.sourcePath.slice(entry.sourcePath.lastIndexOf("/") + 1),
    entry.source === "workbook" ? "xlsx" : "md",
    entry.content,
    entry.embedding,
  );
}
