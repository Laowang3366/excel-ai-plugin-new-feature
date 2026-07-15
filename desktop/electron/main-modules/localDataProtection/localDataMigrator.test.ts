import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { openSqliteDatabase } from "../../agent/storage/nodeSqlite";
import { openOrCreateDataKeystore, type DataKeystoreCipher } from "./dataKeystore";
import { fieldAad } from "./fieldCrypto";
import { migrateLocalDataAtomically } from "./localDataMigrator";
import { createPayloadProtection, setPayloadProtection } from "./payloadProtection";
import { isProtectedBlob } from "./protectedBlob";

const tempDirs: string[] = [];

function makeCipher(): DataKeystoreCipher {
  const map = new Map<string, string>();
  return {
    isAvailable: () => true,
    encrypt: (value) => {
      const token = Buffer.from(value, "utf8").toString("base64");
      map.set(token, value);
      return token;
    },
    decrypt: (value) => {
      const plain = map.get(value);
      if (!plain) throw new Error("missing");
      return plain;
    },
  };
}

afterEach(() => {
  setPayloadProtection(null);
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe("migrateLocalDataAtomically", () => {
  it("stages encrypts and leaves no plaintext canary in SQLite", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-mig-"));
    tempDirs.push(root);
    const dataRoot = path.join(root, "data");
    const userData = path.join(root, "userData");
    fs.mkdirSync(path.join(dataRoot, "sessions", "state-runtime"), { recursive: true });
    fs.mkdirSync(path.join(dataRoot, "knowledge"), { recursive: true });

    const jsonlPath = path.join(dataRoot, "sessions", "2026", "07", "15", "rollout-thread-a.jsonl");
    fs.mkdirSync(path.dirname(jsonlPath), { recursive: true });
    fs.writeFileSync(
      jsonlPath,
      `${JSON.stringify({ timestamp: "2026-07-15T00:00:00.000Z", item: { type: "session_meta", meta: { id: "thread-a", timestamp: "2026-07-15T00:00:00.000Z", modelProvider: "x" } } })}\n`,
      "utf8",
    );

    const logsDbPath = path.join(dataRoot, "sessions", "state-runtime", "logs.db");
    const db = openSqliteDatabase(logsDbPath);
    db.exec(`
      CREATE TABLE rollout_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        thread_id TEXT NOT NULL,
        turn_id TEXT,
        item_type TEXT NOT NULL,
        timestamp TEXT NOT NULL,
        item_json TEXT NOT NULL
      );
      CREATE VIRTUAL TABLE rollout_events_fts USING fts5(
        thread_id UNINDEXED, turn_id UNINDEXED, item_type UNINDEXED, content
      );
    `);
    db.prepare(
      `INSERT INTO rollout_events (thread_id, turn_id, item_type, timestamp, item_json)
       VALUES (?, ?, ?, ?, ?)`,
    ).run(
      "thread-a",
      null,
      "session_meta",
      "t",
      JSON.stringify({ type: "session_meta", secret: "CANARY-PLAINTEXT" }),
    );
    db.close();

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    const protection = createPayloadProtection(keystore);
    setPayloadProtection(protection);

    const handle = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: keystore.currentKeyId,
      previousKeyId: keystore.currentKeyId,
      kind: "encrypt",
      userDataPath: userData,
    });
    if ("finalize" in handle) {
      const finalized = await handle.finalize();
      expect(finalized.cleared).toBe(true);
    }

    const rawJsonl = fs.readFileSync(jsonlPath, "utf8").trim();
    expect(isProtectedBlob(rawJsonl)).toBe(true);
    expect(rawJsonl).not.toContain("session_meta");

    const after = openSqliteDatabase(logsDbPath);
    const row = after.prepare(`SELECT id, item_json FROM rollout_events`).get() as {
      id: number;
      item_json: string;
    };
    expect(isProtectedBlob(row.item_json)).toBe(true);
    expect(row.item_json).not.toContain("CANARY-PLAINTEXT");
    const plain = protection.unprotect(
      row.item_json,
      fieldAad("logs", "rollout_events", String(row.id), "item_json"),
    );
    expect(plain).toContain("CANARY-PLAINTEXT");
    after.close();

    const canaryBuf = Buffer.from("CANARY-PLAINTEXT", "utf8");
    expect(fs.readFileSync(logsDbPath).includes(canaryBuf)).toBe(false);
    for (const suffix of ["-wal", "-shm"]) {
      const side = `${logsDbPath}${suffix}`;
      if (fs.existsSync(side)) {
        expect(fs.readFileSync(side).includes(canaryBuf)).toBe(false);
      }
    }
  });
});
