import type { SqliteDatabase } from "../storage/nodeSqlite";

export function initKnowledgeTables(db: SqliteDatabase): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS knowledge_entries (
      id          TEXT PRIMARY KEY,
      source      TEXT NOT NULL,
      source_path TEXT NOT NULL,
      source_name TEXT NOT NULL,
      source_type TEXT NOT NULL,
      chunk_index INTEGER DEFAULT 0,
      content     TEXT NOT NULL,
      metadata    TEXT DEFAULT '{}',
      embedding   TEXT,
      embedding_provider TEXT,
      embedding_model TEXT,
      embedding_dimensions INTEGER,
      indexed_at  INTEGER NOT NULL,
      token_count INTEGER DEFAULT 0
    );

    CREATE INDEX IF NOT EXISTS idx_entries_source
      ON knowledge_entries(source);
    CREATE INDEX IF NOT EXISTS idx_entries_source_path
      ON knowledge_entries(source_path);
    CREATE INDEX IF NOT EXISTS idx_entries_indexed_at
      ON knowledge_entries(indexed_at);

    CREATE TABLE IF NOT EXISTS knowledge_sources (
      source_path   TEXT PRIMARY KEY,
      source_name   TEXT NOT NULL,
      source_type   TEXT NOT NULL,
      entry_count   INTEGER DEFAULT 0,
      first_indexed INTEGER NOT NULL,
      last_indexed  INTEGER NOT NULL,
      file_hash     TEXT DEFAULT ''
    );
  `);
  migrateEmbeddingProfileColumns(db);
}

function migrateEmbeddingProfileColumns(db: SqliteDatabase): void {
  const rows = db.prepare("PRAGMA table_info(knowledge_entries)").all() as Array<{ name: string }>;
  const columns = new Set(rows.map((row) => row.name));

  if (!columns.has("embedding_provider")) {
    db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_provider TEXT");
  }
  if (!columns.has("embedding_model")) {
    db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_model TEXT");
  }
  if (!columns.has("embedding_dimensions")) {
    db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_dimensions INTEGER");
  }

  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_entries_embedding_profile
      ON knowledge_entries(embedding_provider, embedding_model, embedding_dimensions);
  `);
}
