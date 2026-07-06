import { runSqliteTransaction, type SqliteDatabase } from "../storage/nodeSqlite";

export function ensureSourceSummaries(db: SqliteDatabase): void {
  const rows = db
    .prepare(
      `SELECT
         source_path,
         source_name,
         source_type,
         COUNT(*) as entry_count,
         MIN(indexed_at) as first_indexed,
         MAX(indexed_at) as last_indexed
       FROM knowledge_entries
       GROUP BY source_path, source_name, source_type`
    )
    .all() as Record<string, any>[];

  if (rows.length === 0) return;

  const existingRows = db
    .prepare("SELECT source_path FROM knowledge_sources")
    .all() as Array<{ source_path: string }>;
  const existing = new Set(existingRows.map((row) => row.source_path));

  const insert = db.prepare(
    `INSERT INTO knowledge_sources
      (source_path, source_name, source_type, entry_count,
       first_indexed, last_indexed, file_hash)
     VALUES
      (?, ?, ?, ?,
       ?, ?, ?)`
  );

  runSqliteTransaction(db, () => {
    for (const row of rows) {
      if (existing.has(row.source_path)) continue;
      insert.run(
        row.source_path,
        row.source_name,
        row.source_type,
        row.entry_count,
        row.first_indexed,
        row.last_indexed,
        ""
      );
    }
  });
}
