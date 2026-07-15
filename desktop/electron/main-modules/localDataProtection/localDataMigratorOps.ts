import * as fs from "node:fs";
import * as path from "node:path";
import * as zlib from "node:zlib";
import { promisify } from "node:util";
import { openSqliteDatabase, runSqliteTransaction } from "../../agent/storage/nodeSqlite";
import { fieldAad } from "./fieldCrypto";
import type { PayloadProtection } from "./payloadProtection";
import {
  createRecordId,
  isProtectedBlob,
  jsonlLineAad,
  parseProtectedKeyId,
  parseProtectedRecordId,
} from "./protectedBlob";

const gzipAsync = promisify(zlib.gzip);
const gunzipAsync = promisify(zlib.gunzip);

async function walkFiles(root: string, predicate: (name: string) => boolean): Promise<string[]> {
  const out: string[] = [];
  if (!fs.existsSync(root)) return out;
  const stack = [root];
  while (stack.length > 0) {
    const current = stack.pop()!;
    const entries = await fs.promises.readdir(current, { withFileTypes: true });
    for (const entry of entries) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) stack.push(full);
      else if (entry.isFile() && predicate(entry.name)) out.push(full);
    }
  }
  return out;
}

function sealValue(
  protection: PayloadProtection,
  value: string,
  aad: string,
  targetKeyId: number,
  recordId?: string,
): string {
  if (!value) return value;
  if (isProtectedBlob(value) && parseProtectedKeyId(value) === targetKeyId) return value;
  return protection.protect(
    value,
    aad,
    targetKeyId,
    recordId ?? parseProtectedRecordId(value) ?? undefined,
  );
}

export function sealJsonlText(
  text: string,
  relative: string,
  protection: PayloadProtection,
  targetKeyId: number,
): string {
  const lines = text.split(/\r?\n/);
  const next = lines.map((line) => {
    if (!line.trim()) return line;
    if (isProtectedBlob(line) && parseProtectedKeyId(line) === targetKeyId) return line;
    const rid = parseProtectedRecordId(line) ?? createRecordId();
    const plain = isProtectedBlob(line)
      ? protection.unprotect(line, jsonlLineAad(relative, rid))
      : line;
    return protection.protect(plain, jsonlLineAad(relative, rid), targetKeyId, rid);
  });
  return `${next.join("\n")}${text.endsWith("\n") ? "\n" : ""}`;
}

export async function transformJsonlFiles(
  stageRoot: string,
  protection: PayloadProtection,
  targetKeyId: number,
): Promise<void> {
  const sessionsRoot = path.join(stageRoot, "sessions");
  const files = await walkFiles(
    sessionsRoot,
    (name) => name.endsWith(".jsonl") && !name.includes(".jsonl."),
  );
  for (const filePath of files) {
    const relative = path.relative(sessionsRoot, filePath).split(path.sep).join("/");
    const raw = await fs.promises.readFile(filePath, "utf8");
    if (!raw.trim()) continue;
    const next = sealJsonlText(raw, relative, protection, targetKeyId);
    const tempPath = `${filePath}.${process.pid}.tmp`;
    await fs.promises.writeFile(tempPath, next, "utf8");
    await fs.promises.rename(tempPath, filePath);
  }
}

export async function transformArchiveFiles(
  stageRoot: string,
  protection: PayloadProtection,
  targetKeyId: number,
): Promise<void> {
  const sessionsRoot = path.join(stageRoot, "sessions");
  const gzFiles = await walkFiles(sessionsRoot, (name) => name.endsWith(".jsonl.gz"));
  for (const filePath of gzFiles) {
    const relative = path
      .relative(sessionsRoot, filePath.replace(/\.gz$/u, ""))
      .split(path.sep)
      .join("/");
    const plain = (await gunzipAsync(await fs.promises.readFile(filePath))).toString("utf8");
    const sealed = sealJsonlText(plain, relative, protection, targetKeyId);
    const compressed = await gzipAsync(Buffer.from(sealed, "utf8"));
    const tempPath = `${filePath}.${process.pid}.tmp`;
    await fs.promises.writeFile(tempPath, compressed);
    await fs.promises.rename(tempPath, filePath);
  }

  const zstFiles = await walkFiles(sessionsRoot, (name) => name.endsWith(".jsonl.zst"));
  for (const filePath of zstFiles) {
    const relative = path
      .relative(sessionsRoot, filePath.replace(/\.zst$/u, ""))
      .split(path.sep)
      .join("/");
    // Prefer node:zlib zstd if available; fallback to reading via dynamic require of existing helper.
    const { zstdDecompress, zstdCompress } = await import("./zstdCodec");
    const plain = (await zstdDecompress(await fs.promises.readFile(filePath))).toString("utf8");
    const sealed = sealJsonlText(plain, relative, protection, targetKeyId);
    const compressed = await zstdCompress(Buffer.from(sealed, "utf8"));
    const tempPath = `${filePath}.${process.pid}.tmp`;
    await fs.promises.writeFile(tempPath, compressed);
    await fs.promises.rename(tempPath, filePath);
  }
}

function transformSqliteTable(
  dbPath: string,
  table: string,
  idColumn: string,
  columns: string[],
  store: string,
  protection: PayloadProtection,
  targetKeyId: number,
): void {
  if (!fs.existsSync(dbPath)) return;
  const db = openSqliteDatabase(dbPath);
  try {
    const exists = db
      .prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type = 'table' AND name = ?`)
      .get(table) as { ok?: number } | undefined;
    if (!exists) return;
    const rows = db
      .prepare(`SELECT ${[idColumn, ...columns].join(", ")} FROM ${table}`)
      .all() as Record<string, unknown>[];
    const setters = columns.map((column) => `${column} = ?`).join(", ");
    const update = db.prepare(`UPDATE ${table} SET ${setters} WHERE ${idColumn} = ?`);
    runSqliteTransaction(db, () => {
      for (const row of rows) {
        let changed = false;
        const values: Array<string | number | null> = columns.map((column) => {
          const current = row[column];
          if (typeof current !== "string" || !current) {
            return current == null ? null : String(current);
          }
          const next = sealValue(
            protection,
            current,
            fieldAad(store, table, String(row[idColumn]), column),
            targetKeyId,
          );
          if (next !== current) changed = true;
          return next;
        });
        if (changed) {
          const idValue = row[idColumn];
          update.run(...values, idValue as string | number);
        }
      }
    });
  } finally {
    db.close();
  }
}

function rewriteRolloutFtsEmpty(logsDbPath: string): void {
  if (!fs.existsSync(logsDbPath)) return;
  const db = openSqliteDatabase(logsDbPath);
  try {
    const fts = db
      .prepare(
        `SELECT 1 AS ok FROM sqlite_master WHERE type = 'table' AND name = 'rollout_events_fts'`,
      )
      .get() as { ok?: number } | undefined;
    if (!fts) return;
    db.exec(`DROP TABLE IF EXISTS rollout_events_fts`);
    db.exec(`
      CREATE VIRTUAL TABLE rollout_events_fts USING fts5(
        thread_id UNINDEXED, turn_id UNINDEXED, item_type UNINDEXED, content
      );
    `);
    const rows = db
      .prepare(`SELECT id, thread_id, turn_id, item_type FROM rollout_events`)
      .all() as Array<{ id: number; thread_id: string; turn_id: string | null; item_type: string }>;
    const insert = db.prepare(
      `INSERT INTO rollout_events_fts(rowid, thread_id, turn_id, item_type, content) VALUES (?, ?, ?, ?, ?)`,
    );
    runSqliteTransaction(db, () => {
      for (const row of rows) insert.run(row.id, row.thread_id, row.turn_id, row.item_type, "");
    });
  } finally {
    db.close();
  }
}

export function transformStagePayloads(
  stageRoot: string,
  protection: PayloadProtection,
  targetKeyId: number,
): void {
  const stateRuntime = path.join(stageRoot, "sessions", "state-runtime");
  transformSqliteTable(
    path.join(stateRuntime, "logs.db"),
    "rollout_events",
    "id",
    ["item_json"],
    "logs",
    protection,
    targetKeyId,
  );
  rewriteRolloutFtsEmpty(path.join(stateRuntime, "logs.db"));
  transformSqliteTable(
    path.join(stateRuntime, "logs.db"),
    "tool_execution_logs",
    "id",
    ["arguments_summary", "result_summary", "error", "metadata_json"],
    "logs",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stateRuntime, "memories.db"),
    "memories",
    "memory_id",
    ["content", "metadata_json"],
    "memories",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stateRuntime, "memories.db"),
    "long_term_memories",
    "memory_id",
    ["content", "summary", "metadata_json", "citations_json"],
    "memories",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stateRuntime, "goals.db"),
    "goals",
    "goal_id",
    ["objective", "payload_json"],
    "goals",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stateRuntime, "state.db"),
    "thread_snapshots",
    "thread_id",
    ["preview", "name", "compacted_history"],
    "state",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stateRuntime, "state.db"),
    "thread_names",
    "thread_id",
    ["name"],
    "state",
    protection,
    targetKeyId,
  );
  transformSqliteTable(
    path.join(stageRoot, "knowledge", "knowledge.db"),
    "knowledge_entries",
    "id",
    ["content", "metadata"],
    "knowledge",
    protection,
    targetKeyId,
  );

  // Purge freelist/WAL leftovers that may still hold pre-encryption plaintext pages.
  compactMigratedSqliteDatabases([
    path.join(stateRuntime, "logs.db"),
    path.join(stateRuntime, "memories.db"),
    path.join(stateRuntime, "goals.db"),
    path.join(stateRuntime, "state.db"),
    path.join(stageRoot, "knowledge", "knowledge.db"),
  ]);
}

/** WAL checkpoint + VACUUM for known migrated DBs only — removes old plaintext pages. */
export function compactMigratedSqliteDatabases(dbPaths: string[]): void {
  for (const dbPath of dbPaths) {
    if (!fs.existsSync(dbPath)) continue;
    const db = openSqliteDatabase(dbPath);
    try {
      try {
        db.exec("PRAGMA wal_checkpoint(TRUNCATE)");
      } catch {
        // non-WAL databases may reject checkpoint
      }
      db.exec("VACUUM");
    } finally {
      db.close();
    }
    for (const suffix of ["-wal", "-shm"]) {
      const side = `${dbPath}${suffix}`;
      if (fs.existsSync(side) && fs.statSync(side).size === 0) {
        try {
          fs.rmSync(side, { force: true });
        } catch {
          // leave empty side files if OS holds a handle
        }
      }
    }
  }
}
