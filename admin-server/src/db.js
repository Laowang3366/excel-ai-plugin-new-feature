/**
 * 数据库初始化与连接管理
 *
 * 使用 better-sqlite3（Node.js 兼容性优于 node:sqlite）
 */

import Database from "better-sqlite3";
import path from "path";
import fs from "fs";
import config from "./config.js";

/** 单例数据库实例 */
let db = null;

/**
 * 获取数据库实例（懒初始化）
 */
export function getDb() {
  if (db) return db;

  // 确保数据目录存在
  const dbDir = path.dirname(config.dbPath);
  if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
  }

  db = new Database(config.dbPath);

  // 开启 WAL 模式提升并发性能
  db.pragma("journal_mode = WAL");
  db.pragma("foreign_keys = ON");

  runMigrations(db);

  return db;
}

/**
 * 数据库迁移
 */
function runMigrations(database) {
  database.exec(`
    CREATE TABLE IF NOT EXISTS admin_users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS license_keys (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      key_code TEXT UNIQUE NOT NULL,
      status TEXT NOT NULL DEFAULT 'active',
      duration_days INTEGER,
      max_machines INTEGER DEFAULT 1,
      used_count INTEGER DEFAULT 0,
      note TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      created_by TEXT,
      updated_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS activated_machines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      key_id INTEGER NOT NULL,
      machine_id TEXT NOT NULL,
      machine_name TEXT,
      activated_at TEXT DEFAULT (datetime('now')),
      last_heartbeat TEXT,
      is_online INTEGER DEFAULT 0,
      total_online_seconds INTEGER DEFAULT 0,
      FOREIGN KEY (key_id) REFERENCES license_keys(id),
      UNIQUE(key_id, machine_id)
    );

    CREATE INDEX IF NOT EXISTS idx_license_keys_status ON license_keys(status);
    CREATE INDEX IF NOT EXISTS idx_license_keys_key_code ON license_keys(key_code);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_key_id ON activated_machines(key_id);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_machine_id ON activated_machines(machine_id);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_is_online ON activated_machines(is_online);
  `);
}

/**
 * 关闭数据库连接
 */
export function closeDb() {
  if (db) {
    db.close();
    db = null;
  }
}
