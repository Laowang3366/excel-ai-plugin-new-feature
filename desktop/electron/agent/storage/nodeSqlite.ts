import { DatabaseSync } from "node:sqlite";

export type SqliteDatabase = InstanceType<typeof DatabaseSync>;

export function openSqliteDatabase(dbPath: string): SqliteDatabase {
  return new DatabaseSync(dbPath);
}

export function runPragma(db: SqliteDatabase, pragma: string): unknown {
  return db.prepare(`PRAGMA ${pragma}`).get();
}

export function runPragmaValue(db: SqliteDatabase, pragma: string): string {
  const row = runPragma(db, pragma);
  if (row && typeof row === "object") {
    const values = Object.values(row as Record<string, unknown>);
    if (values.length > 0) return String(values[0]);
  }
  return String(row ?? "");
}

export function runSqliteTransaction<T>(db: SqliteDatabase, fn: () => T): T {
  db.prepare("BEGIN").run();
  try {
    const result = fn();
    db.prepare("COMMIT").run();
    return result;
  } catch (error) {
    try {
      db.prepare("ROLLBACK").run();
    } catch {
      // SQLite may already have closed the transaction after a failed COMMIT.
    }
    throw error;
  }
}
