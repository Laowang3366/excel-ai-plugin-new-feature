/**
 * 首次安装脚本 — 创建默认管理员账号
 *
 * 用法：node setup.js
 *
 * 可选环境变量覆盖默认凭据：
 *   ADMIN_USERNAME=admin
 *   ADMIN_PASSWORD=admin123
 *
 * 工作流程：
 * 1. 初始化数据库（调用 src/db.js 的 getDb，自动建表）
 * 2. 检查 admin_users 表是否存在管理员
 *    - 存在则显示已有管理员列表，跳过创建
 *    - 不存在则使用 bcrypt 哈希密码并插入新管理员
 * 3. 输出数据库统计信息（卡密数、设备数）
 * 4. 打印登录地址和凭据
 *
 * 安全性提醒：
 * - 默认密码 admin123 仅适用于开发环境
 * - 生产部署后务必通过环境变量 ADMIN_PASSWORD 设置强密码
 * - 或登录后台后手动修改密码（目前无修改密码接口，需直接操作数据库）
 */

import bcrypt from "bcryptjs";
import { getDb, closeDb } from "./src/db.js";

// 从环境变量读取管理员凭据，未设置则使用开发默认值
const username = process.env.ADMIN_USERNAME || "admin";
const password = process.env.ADMIN_PASSWORD || "admin123";

async function main() {
  console.log("========================================");
  console.log("  文格 AI 助手 — 激活管理后台安装脚本");
  console.log("========================================\n");

  // 初始化数据库连接（自动执行迁移建表）
  const db = getDb();

  // 检查是否已存在管理员账号
  const existingAdmin = db.prepare("SELECT COUNT(*) as count FROM admin_users").get();

  if (existingAdmin.count > 0) {
    // 已有管理员，跳过创建并列出已有账号信息
    console.log("⚠️  数据库已存在管理员账号，跳过创建。\n");
    const users = db.prepare("SELECT id, username, created_at FROM admin_users").all();
    console.log("现有管理员账号：");
    users.forEach((u) => {
      console.log(`  - ${u.username} (创建于 ${u.created_at})`);
    });
    console.log("");
  } else {
    // 首次安装：生成密码哈希并插入管理员记录
    const salt = bcrypt.genSaltSync(10);
    const hash = bcrypt.hashSync(password, salt);

    db.prepare("INSERT INTO admin_users (username, password_hash) VALUES (?, ?)").run(username, hash);

    console.log("✅ 管理员账号创建成功！\n");
  }

  // 输出当前数据库的数据量统计
  const keyCount = db.prepare("SELECT COUNT(*) as count FROM license_keys").get();
  const machineCount = db.prepare("SELECT COUNT(*) as count FROM activated_machines").get();

  console.log(`数据库状态：
  - 管理员账号: ${existingAdmin.count > 0 ? existingAdmin.count : 1} 个
  - 卡密: ${keyCount.count} 个
  - 已激活设备: ${machineCount.count} 台\n`);

  console.log("========================================");
  console.log("  管理后台登录信息");
  console.log("========================================");
  console.log(`  地址: http://localhost:3456`);
  console.log(`  用户名: ${username}`);
  console.log(`  密码: ${password}`);
  console.log("========================================\n");
  console.log("⚠️  请及时修改默认密码！\n");

  // 清理数据库连接
  closeDb();
}

main().catch((err) => {
  console.error("安装失败:", err);
  process.exit(1);
});
