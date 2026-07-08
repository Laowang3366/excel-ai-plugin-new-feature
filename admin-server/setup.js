/**
 * 首次安装脚本
 *
 * 创建默认管理员账号并输出提示信息。
 *
 * 用法：node setup.js
 * 可选环境变量：
 *   ADMIN_USERNAME=admin
 *   ADMIN_PASSWORD=admin123
 */

import bcrypt from "bcryptjs";
import { getDb, closeDb } from "./src/db.js";

const username = process.env.ADMIN_USERNAME || "admin";
const password = process.env.ADMIN_PASSWORD || "admin123";

async function main() {
  console.log("========================================");
  console.log("  文格 AI 助手 — 激活管理后台安装脚本");
  console.log("========================================\n");

  const db = getDb();

  // 检查是否已有管理员
  const existingAdmin = db.prepare("SELECT COUNT(*) as count FROM admin_users").get();

  if (existingAdmin.count > 0) {
    console.log("⚠️  数据库已存在管理员账号，跳过创建。\n");
    const users = db.prepare("SELECT id, username, created_at FROM admin_users").all();
    console.log("现有管理员账号：");
    users.forEach((u) => {
      console.log(`  - ${u.username} (创建于 ${u.created_at})`);
    });
    console.log("");
  } else {
    // 创建管理员
    const salt = bcrypt.genSaltSync(10);
    const hash = bcrypt.hashSync(password, salt);

    db.prepare("INSERT INTO admin_users (username, password_hash) VALUES (?, ?)").run(username, hash);

    console.log("✅ 管理员账号创建成功！\n");
  }

  // 输出卡密统计
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

  closeDb();
}

main().catch((err) => {
  console.error("安装失败:", err);
  process.exit(1);
});
