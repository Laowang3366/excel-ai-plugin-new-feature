import { createHmac } from "node:crypto";
import fs from "node:fs";
import path from "node:path";

import Database from "better-sqlite3";

const DAY_MS = 24 * 60 * 60 * 1000;
const DEFAULT_CLEANUP_INTERVAL_MS = 6 * 60 * 60 * 1000;
const ANALYTICS_SCHEMA_VERSION = 1;

function normalizeReferer(value) {
  try {
    const url = new URL(String(value || ""));
    return url.protocol === "http:" || url.protocol === "https:" ? url.origin.slice(0, 200) : "";
  } catch {
    return "";
  }
}

function migrateAnalyticsData(db) {
  if (db.pragma("user_version", { simple: true }) >= ANALYTICS_SCHEMA_VERSION) return;
  const selectReferers = db.prepare("SELECT id, referer FROM downloads WHERE referer <> ''");
  const updateReferer = db.prepare("UPDATE downloads SET referer = ? WHERE id = ?");
  const migrate = db.transaction(() => {
    db.exec("UPDATE downloads SET user_agent = substr(user_agent, 1, 200)");
    for (const row of selectReferers.all()) {
      updateReferer.run(normalizeReferer(row.referer), row.id);
    }
    db.pragma(`user_version = ${ANALYTICS_SCHEMA_VERSION}`);
  });
  migrate();
}

export function createAnalyticsDatabase(databasePath, options) {
  const {
    analyticsSalt,
    retentionDays = 90,
    ipRotationDays = 30,
    now = Date.now,
    cleanupIntervalMs = DEFAULT_CLEANUP_INTERVAL_MS,
    onMaintenanceError = () => {},
  } = options;
  fs.mkdirSync(path.dirname(databasePath), { recursive: true });
  const db = new Database(databasePath);
  db.pragma("journal_mode = WAL");
  db.pragma("busy_timeout = 5000");
  db.exec(`
    CREATE TABLE IF NOT EXISTS downloads (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      downloaded_at INTEGER NOT NULL,
      version TEXT NOT NULL,
      artifact TEXT NOT NULL,
      ip_hash TEXT NOT NULL,
      user_agent TEXT NOT NULL,
      referer TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_downloads_time ON downloads(downloaded_at);
    CREATE INDEX IF NOT EXISTS idx_downloads_version ON downloads(version);
  `);

  const deleteExpired = db.prepare("DELETE FROM downloads WHERE downloaded_at < ?");

  function cleanupExpired() {
    try {
      const cutoff = now() - retentionDays * DAY_MS;
      return deleteExpired.run(cutoff).changes;
    } catch (error) {
      try {
        onMaintenanceError(error);
      } catch {
        // Maintenance and its reporting are both best-effort.
      }
      return 0;
    }
  }

  cleanupExpired();
  migrateAnalyticsData(db);

  const insertDownload = db.prepare(`
    INSERT INTO downloads (downloaded_at, version, artifact, ip_hash, user_agent, referer)
    VALUES (@downloadedAt, @version, @artifact, @ipHash, @userAgent, @referer)
  `);

  function hashIp(ip, downloadedAt) {
    const period = Math.floor(downloadedAt / (ipRotationDays * DAY_MS));
    const periodKey = createHmac("sha256", analyticsSalt).update(`analytics-ip:${period}`).digest();
    return createHmac("sha256", periodKey).update(ip || "unknown").digest("hex");
  }

  const cleanupTimer = setInterval(cleanupExpired, cleanupIntervalMs);
  cleanupTimer.unref?.();

  return {
    recordDownload({ version, artifact, ip, userAgent, referer }) {
      const downloadedAt = now();
      insertDownload.run({
        downloadedAt,
        version,
        artifact,
        ipHash: hashIp(ip, downloadedAt),
        userAgent: String(userAgent || "").slice(0, 200),
        referer: normalizeReferer(referer),
      });
    },
    getStats(days = 30) {
      const safeDays = Math.max(1, Math.min(365, Number(days) || 30));
      const currentTime = now();
      const start = currentTime - safeDays * DAY_MS;
      const todayStart = Math.floor(currentTime / DAY_MS) * DAY_MS;
      const summary = db.prepare(`
        SELECT
          COUNT(*) AS total,
          COUNT(DISTINCT ip_hash) AS uniqueDownloads,
          SUM(CASE WHEN downloaded_at >= ? THEN 1 ELSE 0 END) AS today
        FROM downloads WHERE downloaded_at >= ?
      `).get(todayStart, start);
      const daily = db.prepare(`
        SELECT strftime('%Y-%m-%d', downloaded_at / 1000, 'unixepoch') AS day, COUNT(*) AS downloads
        FROM downloads WHERE downloaded_at >= ? GROUP BY day ORDER BY day ASC
      `).all(start);
      const versions = db.prepare(`
        SELECT version, COUNT(*) AS downloads
        FROM downloads WHERE downloaded_at >= ? GROUP BY version ORDER BY downloads DESC
      `).all(start);
      const recent = db.prepare(`
        SELECT downloaded_at AS downloadedAt, version, artifact,
               substr(ip_hash, 1, 12) AS visitor, user_agent AS userAgent, referer
        FROM downloads ORDER BY downloaded_at DESC LIMIT 50
      `).all();
      return { days: safeDays, summary, daily, versions, recent };
    },
    close() {
      clearInterval(cleanupTimer);
      db.close();
    },
  };
}
