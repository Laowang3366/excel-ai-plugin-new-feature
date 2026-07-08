/**
 * 数据库初始化与连接管理
 *
 * 使用 better-sqlite3（Node.js 兼容性优于 node:sqlite）。
 * 采用单例模式：整个进程共享一个数据库连接。
 *
 * 迁移策略说明：
 * - 使用 CREATE TABLE IF NOT EXISTS，首次启动自动建表。
 * - 若后续需要 schema 变更，应在此处增加 ALTER TABLE 或版本号判断逻辑，
 *   而非手动改库，避免生产环境遗漏迁移。
 *
 * WAL 模式说明：
 * - WAL（Write-Ahead Logging）允许并发读取，写入不阻塞读取，
 *   对于本系统少量写入 + 大量查询的场景（如后台浏览卡密列表）有显著性能优势。
 * - foreign_keys = ON 确保 activated_machines.key_id 引用完整性。
 */

import Database from "better-sqlite3";
import path from "path";
import fs from "fs";
import config from "./config.js";

/** 单例数据库实例（模块级缓存） */
let db = null;

/**
 * 获取数据库实例（懒初始化）
 *
 * 首次调用时自动创建数据目录、打开数据库、启用 WAL 模式并执行迁移。
 * 后续调用直接返回缓存实例。
 */
export function getDb() {
  if (db) return db;

  // 确保数据目录存在，避免 better-sqlite3 因目录不存在而抛错
  const dbDir = path.dirname(config.dbPath);
  if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
  }

  db = new Database(config.dbPath);

  // 开启 WAL 模式以提升并发读性能
  db.pragma("journal_mode = WAL");
  // 启用外键约束保证数据引用完整性
  db.pragma("foreign_keys = ON");

  runMigrations(db);

  return db;
}

/**
 * 数据库迁移（幂等执行）
 *
 * 所有 DDL 使用 IF NOT EXISTS 确保重复执行安全。
 *
 * === 表结构说明 ===
 *
 * admin_users        — 管理员账号，密码以 bcrypt 哈希存储
 * license_keys       — 卡密主表，记录卡密、状态、有效期、绑定上限等
 * activated_machines — 设备绑定表，记录每张卡密下的激活设备及在线状态
 *
 * === 索引设计 ===
 * license_keys:
 *   - idx_license_keys_status   — 按状态筛选（仪表盘统计、列表过滤）
 *   - idx_license_keys_key_code — 按卡密精确查找（激活验证首要查询路径）
 * activated_machines:
 *   - idx_activated_machines_key_id      — 按卡密查询关联设备
 *   - idx_activated_machines_machine_id  — 按设备 ID 查询（心跳上报）
 *   - idx_activated_machines_is_online   — 在线设备统计（仪表盘）
 */
function runMigrations(database) {
  database.exec(`
    -- 管理员账号表
    CREATE TABLE IF NOT EXISTS admin_users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,       -- 登录名，唯一约束
      password_hash TEXT NOT NULL,          -- bcrypt 哈希值
      created_at TEXT DEFAULT (datetime('now'))
    );

    -- 卡密表
    CREATE TABLE IF NOT EXISTS license_keys (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      key_code TEXT UNIQUE NOT NULL,        -- 卡密字符串（大写字母+数字）
      status TEXT NOT NULL DEFAULT 'active', -- active | disabled | used | expired
      duration_days INTEGER,                -- 有效期天数，NULL 表示永久
      max_machines INTEGER DEFAULT 1,       -- 最大绑定设备数
      used_count INTEGER DEFAULT 0,         -- 已绑定设备计数
      note TEXT,                            -- 管理员备注
      created_at TEXT DEFAULT (datetime('now')),
      created_by TEXT,                      -- 创建者用户名
      updated_at TEXT DEFAULT (datetime('now'))
    );

    -- 设备绑定表
    CREATE TABLE IF NOT EXISTS activated_machines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      key_id INTEGER NOT NULL,              -- 关联卡密 ID
      machine_id TEXT NOT NULL,             -- 客户端设备唯一标识
      machine_name TEXT,                    -- 设备友好名称
      activated_at TEXT DEFAULT (datetime('now')),
      last_heartbeat TEXT,                  -- 最近一次心跳时间
      is_online INTEGER DEFAULT 0,          -- 在线标志（冗余字段，便于快速查询）
      total_online_seconds INTEGER DEFAULT 0, -- 累计在线时长（秒）
      FOREIGN KEY (key_id) REFERENCES license_keys(id) ON DELETE CASCADE,
      UNIQUE(key_id, machine_id)           -- 同一卡密下设备 ID 唯一
    );

    -- 查询加速索引
    CREATE INDEX IF NOT EXISTS idx_license_keys_status ON license_keys(status);
    CREATE INDEX IF NOT EXISTS idx_license_keys_key_code ON license_keys(key_code);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_key_id ON activated_machines(key_id);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_machine_id ON activated_machines(machine_id);
    CREATE INDEX IF NOT EXISTS idx_activated_machines_is_online ON activated_machines(is_online);
  `);
}

/**
 * 关闭数据库连接（通常在应用退出时调用）
 */
export function closeDb() {
  if (db) {
    db.close();
    db = null;
  }
}
