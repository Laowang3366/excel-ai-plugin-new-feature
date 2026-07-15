import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import Database from "better-sqlite3";

import { createAnalyticsDatabase } from "../src/database.mjs";

const DAY_MS = 24 * 60 * 60 * 1000;
const LONG_CLEANUP_INTERVAL_MS = 1_000_000_000;

async function createTemporaryDatabase(prefix) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), prefix));
  return { root, databasePath: path.join(root, "analytics.sqlite") };
}

test("IP identifiers are stable within a rotation period and change across periods", async (context) => {
  const { root, databasePath } = await createTemporaryDatabase("wenge-analytics-rotation-");
  let currentTime = 100 * 30 * DAY_MS;
  const analytics = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 365,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
  });
  context.after(async () => {
    analytics.close();
    await fs.rm(root, { recursive: true, force: true });
  });

  analytics.recordDownload({ version: "1.0.0", artifact: "setup.exe", ip: "198.51.100.8" });
  currentTime += DAY_MS;
  analytics.recordDownload({ version: "1.0.0", artifact: "setup.exe", ip: "198.51.100.8" });
  currentTime += 30 * DAY_MS;
  analytics.recordDownload({ version: "1.0.0", artifact: "setup.exe", ip: "198.51.100.8" });

  const visitors = analytics.getStats(365).recent.map((row) => row.visitor);
  assert.equal(visitors[1], visitors[2]);
  assert.notEqual(visitors[0], visitors[1]);
});

test("startup cleanup deletes expired downloads and retains rows inside the retention window", async (context) => {
  const { root, databasePath } = await createTemporaryDatabase("wenge-analytics-retention-");
  let currentTime = 0;
  const initial = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 90,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
  });
  initial.recordDownload({ version: "old", artifact: "old.exe", ip: "198.51.100.1" });
  currentTime = 81 * DAY_MS;
  initial.recordDownload({ version: "current", artifact: "current.exe", ip: "198.51.100.2" });
  initial.close();

  currentTime = 100 * DAY_MS;
  const reopened = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 90,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
  });
  context.after(async () => {
    reopened.close();
    await fs.rm(root, { recursive: true, force: true });
  });

  const stats = reopened.getStats(365);
  assert.equal(stats.summary.total, 1);
  assert.equal(stats.recent[0].version, "current");
});

test("legacy analytics rows are minimized during migration", async (context) => {
  const { root, databasePath } = await createTemporaryDatabase("wenge-analytics-migration-");
  const legacy = new Database(databasePath);
  legacy.exec(`
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
  legacy.prepare(`
    INSERT INTO downloads (downloaded_at, version, artifact, ip_hash, user_agent, referer)
    VALUES (?, ?, ?, ?, ?, ?)
  `).run(DAY_MS, "legacy", "legacy.exe", "a".repeat(64), "u".repeat(500), "https://example.test/private/path?token=secret");
  legacy.close();

  let currentTime = 2 * DAY_MS;
  const analytics = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 90,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
  });
  context.after(async () => {
    analytics.close();
    await fs.rm(root, { recursive: true, force: true });
  });

  let recent = analytics.getStats(30).recent;
  assert.equal(recent[0].userAgent.length, 200);
  assert.equal(recent[0].referer, "https://example.test");

  currentTime += DAY_MS;
  analytics.recordDownload({
    version: "current",
    artifact: "current.exe",
    ip: "198.51.100.3",
    userAgent: "n".repeat(500),
    referer: "https://downloads.example.test/campaign?customer=private",
  });
  recent = analytics.getStats(30).recent;
  assert.equal(recent[0].userAgent.length, 200);
  assert.equal(recent[0].referer, "https://downloads.example.test");
});

test("cleanup failures are reported without preventing new analytics writes", async (context) => {
  const { root, databasePath } = await createTemporaryDatabase("wenge-analytics-cleanup-failure-");
  let currentTime = 0;
  const initial = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 30,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
  });
  initial.recordDownload({ version: "old", artifact: "old.exe", ip: "198.51.100.4" });
  initial.close();

  const raw = new Database(databasePath);
  raw.exec(`
    CREATE TRIGGER fail_download_cleanup
    BEFORE DELETE ON downloads
    BEGIN
      SELECT RAISE(ABORT, 'cleanup failure');
    END
  `);
  raw.close();

  currentTime = 100 * DAY_MS;
  const maintenanceErrors = [];
  const analytics = createAnalyticsDatabase(databasePath, {
    analyticsSalt: "test-analytics-salt",
    retentionDays: 30,
    ipRotationDays: 30,
    cleanupIntervalMs: LONG_CLEANUP_INTERVAL_MS,
    now: () => currentTime,
    onMaintenanceError(error) {
      maintenanceErrors.push(error);
      throw new Error("monitor unavailable");
    },
  });
  context.after(async () => {
    analytics.close();
    await fs.rm(root, { recursive: true, force: true });
  });

  analytics.recordDownload({ version: "current", artifact: "current.exe", ip: "198.51.100.5" });
  assert.equal(maintenanceErrors.length, 1);
  assert.equal(analytics.getStats(365).summary.total, 2);
});
