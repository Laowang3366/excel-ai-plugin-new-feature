import { createHmac } from "node:crypto";
import fs from "node:fs";
import path from "node:path";

import Database from "better-sqlite3";

export function createAnalyticsDatabase(databasePath, analyticsSalt) {
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

  const insertDownload = db.prepare(`
    INSERT INTO downloads (downloaded_at, version, artifact, ip_hash, user_agent, referer)
    VALUES (@downloadedAt, @version, @artifact, @ipHash, @userAgent, @referer)
  `);

  function hashIp(ip) {
    return createHmac("sha256", analyticsSalt).update(ip || "unknown").digest("hex");
  }

  return {
    recordDownload({ version, artifact, ip, userAgent, referer }) {
      insertDownload.run({
        downloadedAt: Date.now(),
        version,
        artifact,
        ipHash: hashIp(ip),
        userAgent: String(userAgent || "").slice(0, 500),
        referer: String(referer || "").slice(0, 500),
      });
    },
    getStats(days = 30) {
      const safeDays = Math.max(1, Math.min(365, Number(days) || 30));
      const now = Date.now();
      const start = now - safeDays * 24 * 60 * 60 * 1000;
      const todayStart = new Date(new Date().toDateString()).getTime();
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
      db.close();
    },
  };
}
