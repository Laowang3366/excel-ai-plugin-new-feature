/**
 * 文格 AI 助手 - 激活管理后台
 *
 * Express 服务器入口
 */

import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import config from "./config.js";
import { getDb } from "./db.js";
import { jsonBodyParser } from "./middleware/jsonBody.js";
import adminRouter from "./routes/admin.js";
import activateRouter from "./routes/activate.js";
import heartbeatRouter from "./routes/heartbeat.js";
import devicesRouter from "./routes/devices.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();

// ============================================================
// 中间件
// ============================================================

app.use(cors());
app.use(jsonBodyParser());
app.use(express.urlencoded({ extended: true }));

// 静态文件（管理 UI）
app.use(express.static(path.join(__dirname, "public")));

// ============================================================
// API 路由
// ============================================================

app.use("/api/admin", adminRouter);
app.use("/api/activate", activateRouter);
app.use("/api/heartbeat", heartbeatRouter);
app.use("/api/devices", devicesRouter);

// ============================================================
// SPA 回退：所有非 API 请求返回 index.html
// ============================================================

app.get("*", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ============================================================
// 启动服务器
// ============================================================

function start() {
  // 初始化数据库
  getDb();
  console.log(`[激活后台] 数据库已连接: ${config.dbPath}`);

  app.listen(config.port, () => {
    console.log(`[激活后台] 管理后台已启动: http://localhost:${config.port}`);
    console.log(`[激活后台] 首次使用请运行: npm run setup`);
  });
}

start();
