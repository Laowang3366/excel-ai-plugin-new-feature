import * as fs from "node:fs";
import * as path from "node:path";
import * as zlib from "node:zlib";
import { promisify } from "node:util";
import { SqliteStore } from "../../agent/knowledge";
import { StateRuntimeStore } from "../../agent/memory/stateRuntimeStore";
import { openSqliteDatabase } from "../../agent/storage/nodeSqlite";
import { fieldAad } from "./fieldCrypto";
import type { PayloadProtection } from "./payloadProtection";
import {
  isProtectedBlob,
  jsonlLineAad,
  parseProtectedKeyId,
  parseProtectedRecordId,
} from "./protectedBlob";

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

function assertSealedColumn(
  value: unknown,
  aad: string,
  protection: PayloadProtection,
  targetKeyId: number,
  label: string,
): void {
  if (typeof value !== "string" || !value) return;
  if (!isProtectedBlob(value) || parseProtectedKeyId(value) !== targetKeyId) {
    throw new Error(`migration validation failed: ${label} not sealed with target key`);
  }
  protection.unprotect(value, aad);
}

function validateJsonlText(
  text: string,
  relative: string,
  protection: PayloadProtection,
  targetKeyId: number,
  label: string,
): void {
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) continue;
    const rid = parseProtectedRecordId(line);
    if (!rid || parseProtectedKeyId(line) !== targetKeyId) {
      throw new Error(`migration validation failed: ${label} ${relative}`);
    }
    protection.unprotect(line, jsonlLineAad(relative, rid));
  }
}

export async function validateAllEncrypted(
  stageRoot: string,
  protection: PayloadProtection,
  targetKeyId: number,
): Promise<void> {
  const sessionsRoot = path.join(stageRoot, "sessions");
  const jsonlFiles = await walkFiles(
    sessionsRoot,
    (name) => name.endsWith(".jsonl") && !name.includes(".jsonl."),
  );
  for (const filePath of jsonlFiles) {
    const relative = path.relative(sessionsRoot, filePath).split(path.sep).join("/");
    const raw = await fs.promises.readFile(filePath, "utf8");
    validateJsonlText(raw, relative, protection, targetKeyId, "jsonl");
  }

  const gzFiles = await walkFiles(sessionsRoot, (name) => name.endsWith(".jsonl.gz"));
  for (const filePath of gzFiles) {
    const relative = path
      .relative(sessionsRoot, filePath.replace(/\.gz$/u, ""))
      .split(path.sep)
      .join("/");
    const plain = (await gunzipAsync(await fs.promises.readFile(filePath))).toString("utf8");
    validateJsonlText(plain, relative, protection, targetKeyId, "jsonl.gz");
  }

  const zstFiles = await walkFiles(sessionsRoot, (name) => name.endsWith(".jsonl.zst"));
  for (const filePath of zstFiles) {
    const relative = path
      .relative(sessionsRoot, filePath.replace(/\.zst$/u, ""))
      .split(path.sep)
      .join("/");
    const { zstdDecompress } = await import("./zstdCodec");
    const plain = (await zstdDecompress(await fs.promises.readFile(filePath))).toString("utf8");
    validateJsonlText(plain, relative, protection, targetKeyId, "jsonl.zst");
  }

  const stateRuntime = path.join(stageRoot, "sessions", "state-runtime");
  const checks: Array<{ db: string; table: string; id: string; columns: string[]; store: string }> =
    [
      {
        db: path.join(stateRuntime, "logs.db"),
        table: "rollout_events",
        id: "id",
        columns: ["item_json"],
        store: "logs",
      },
      {
        db: path.join(stateRuntime, "logs.db"),
        table: "tool_execution_logs",
        id: "id",
        columns: ["arguments_summary", "result_summary", "error", "metadata_json"],
        store: "logs",
      },
      {
        db: path.join(stateRuntime, "memories.db"),
        table: "memories",
        id: "memory_id",
        columns: ["content", "metadata_json"],
        store: "memories",
      },
      {
        db: path.join(stateRuntime, "memories.db"),
        table: "long_term_memories",
        id: "memory_id",
        columns: ["content", "summary", "metadata_json", "citations_json"],
        store: "memories",
      },
      {
        db: path.join(stateRuntime, "goals.db"),
        table: "goals",
        id: "goal_id",
        columns: ["objective", "payload_json"],
        store: "goals",
      },
      {
        db: path.join(stateRuntime, "state.db"),
        table: "thread_snapshots",
        id: "thread_id",
        columns: ["preview", "name", "compacted_history"],
        store: "state",
      },
      {
        db: path.join(stateRuntime, "state.db"),
        table: "thread_names",
        id: "thread_id",
        columns: ["name"],
        store: "state",
      },
      {
        db: path.join(stageRoot, "knowledge", "knowledge.db"),
        table: "knowledge_entries",
        id: "id",
        columns: ["content", "metadata"],
        store: "knowledge",
      },
    ];
  for (const check of checks) {
    if (!fs.existsSync(check.db)) continue;
    const db = openSqliteDatabase(check.db);
    try {
      const exists = db
        .prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type = 'table' AND name = ?`)
        .get(check.table) as { ok?: number } | undefined;
      if (!exists) continue;
      const rows = db
        .prepare(`SELECT ${[check.id, ...check.columns].join(", ")} FROM ${check.table}`)
        .all() as Record<string, unknown>[];
      for (const row of rows) {
        for (const column of check.columns) {
          assertSealedColumn(
            row[column],
            fieldAad(check.store, check.table, String(row[check.id]), column),
            protection,
            targetKeyId,
            `${check.table}.${column}`,
          );
        }
      }
    } finally {
      db.close();
    }
  }

  const stateStore = new StateRuntimeStore(path.join(stageRoot, "sessions", "state-runtime"));
  await stateStore.init();
  await stateStore.close();
  const knowledgePath = path.join(stageRoot, "knowledge", "knowledge.db");
  if (fs.existsSync(knowledgePath)) {
    const knowledgeStore = new SqliteStore(knowledgePath);
    try {
      knowledgeStore.init();
    } finally {
      knowledgeStore.close();
    }
  }
}
