import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import Database from "better-sqlite3";

import {
  createAnalyticsBackup,
  restoreAnalyticsBackup,
  verifyAnalyticsBackup,
} from "../src/analyticsBackup.mjs";

function openAnalyticsDatabase(databasePath) {
  const db = new Database(databasePath);
  db.pragma("journal_mode = WAL");
  db.exec(`
    CREATE TABLE downloads (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      downloaded_at INTEGER NOT NULL,
      version TEXT NOT NULL,
      artifact TEXT NOT NULL,
      ip_hash TEXT NOT NULL,
      user_agent TEXT NOT NULL,
      referer TEXT NOT NULL
    )
  `);
  return db;
}

function insertDownload(db, version) {
  db.prepare(
    `
    INSERT INTO downloads (downloaded_at, version, artifact, ip_hash, user_agent, referer)
    VALUES (?, ?, ?, ?, '', '')
  `,
  ).run(Date.now(), version, `${version}.exe`, version.padEnd(64, "0"));
}

async function createTemporaryRoot(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

test("online backups restore a consistent snapshot without later writes", async (context) => {
  const root = await createTemporaryRoot("wenge-backup-restore-");
  const sourcePath = path.join(root, "data", "analytics.sqlite");
  const outputDir = path.join(root, "backups");
  const targetPath = path.join(root, "restore", "analytics.sqlite");
  await fs.mkdir(path.dirname(sourcePath), { recursive: true });
  const writer = openAnalyticsDatabase(sourcePath);
  context.after(async () => {
    if (writer.open) writer.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  insertDownload(writer, "before-backup");

  const backup = await createAnalyticsBackup({
    sourcePath,
    outputDir,
    now: () => Date.UTC(2026, 6, 15),
  });
  insertDownload(writer, "after-backup");
  await verifyAnalyticsBackup(backup.backupPath);
  await restoreAnalyticsBackup({ backupPath: backup.backupPath, targetPath });

  const restored = new Database(targetPath, { readonly: true });
  try {
    assert.deepEqual(
      restored.prepare("SELECT version FROM downloads ORDER BY id").all(),
      [{ version: "before-backup" }],
    );
  } finally {
    restored.close();
  }
});

test("backup verification rejects tampering even when metadata is rewritten", async (context) => {
  const root = await createTemporaryRoot("wenge-backup-corrupt-");
  context.after(async () => fs.rm(root, { recursive: true, force: true }));
  const sourcePath = path.join(root, "analytics.sqlite");
  const writer = openAnalyticsDatabase(sourcePath);
  insertDownload(writer, "1.0.0");
  writer.close();
  const backup = await createAnalyticsBackup({
    sourcePath,
    outputDir: path.join(root, "backups"),
  });

  const bytes = await fs.readFile(backup.backupPath);
  bytes.fill(0, 0, 16);
  await fs.writeFile(backup.backupPath, bytes);
  const metadata = JSON.parse(await fs.readFile(backup.metadataPath, "utf8"));
  metadata.sha256 = createHash("sha256").update(bytes).digest("hex");
  metadata.size = bytes.length;
  await fs.writeFile(
    backup.metadataPath,
    `${JSON.stringify(metadata, null, 2)}\n`,
  );

  await assert.rejects(() => verifyAnalyticsBackup(backup.backupPath));
});

test("backup retention removes only the oldest complete backup pairs", async (context) => {
  const root = await createTemporaryRoot("wenge-backup-retention-");
  context.after(async () => fs.rm(root, { recursive: true, force: true }));
  const sourcePath = path.join(root, "analytics.sqlite");
  const outputDir = path.join(root, "backups");
  const writer = openAnalyticsDatabase(sourcePath);
  insertDownload(writer, "1.0.0");
  writer.close();

  for (let day = 1; day <= 3; day += 1) {
    await createAnalyticsBackup({
      sourcePath,
      outputDir,
      retain: 2,
      now: () => Date.UTC(2026, 6, day),
    });
  }

  const files = await fs.readdir(outputDir);
  assert.equal(files.filter((file) => file.endsWith(".sqlite")).length, 2);
  assert.equal(files.filter((file) => file.endsWith(".sqlite.json")).length, 2);
  assert.equal(
    files.some((file) => file.includes("20260701")),
    false,
    JSON.stringify(files),
  );
});

test("restore refuses to overwrite an existing database", async (context) => {
  const root = await createTemporaryRoot("wenge-backup-existing-");
  context.after(async () => fs.rm(root, { recursive: true, force: true }));
  const sourcePath = path.join(root, "analytics.sqlite");
  const writer = openAnalyticsDatabase(sourcePath);
  writer.close();
  const backup = await createAnalyticsBackup({
    sourcePath,
    outputDir: path.join(root, "backups"),
  });
  const targetPath = path.join(root, "target.sqlite");
  await fs.writeFile(targetPath, "do-not-overwrite");

  await assert.rejects(
    () => restoreAnalyticsBackup({ backupPath: backup.backupPath, targetPath }),
    /恢复目标已存在/,
  );
  assert.equal(await fs.readFile(targetPath, "utf8"), "do-not-overwrite");
});
