/**
 * 文格 AI 助手 — 激活管理后台 Express 服务器入口
 *
 * 职责：
 * 1. 组装中间件管线（CORS、自定义 JSON 解析、URL-encoded 解析、静态文件）
 * 2. 挂载 API 路由（管理端、激活端、心跳端、设备端）
 * 3. 提供 SPA 回退路由，确保前端路由刷新时仍返回 index.html
 * 4. 启动服务器并初始化数据库
 *
 * 中间件执行顺序说明：
 * Express 按 app.use() 注册顺序依次执行中间件。
 * CORS 必须排在最前以处理跨域预检请求（OPTIONS）；
 * 自定义 jsonBodyParser 替代 express.json() 以兼容 GBK 编码的请求体；
 * 静态文件中间件在 API 路由之前，但实际路径不冲突（/api/* 不匹配静态文件）；
 * SPA 回退（catch-all）放在最后，确保 API 路由优先匹配。
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
// 全局中间件（按顺序执行）
// ============================================================

// 1. CORS — 允许跨域请求（开发环境可能前后端分离）
app.use(cors());

// 2. JSON 解析 — 替代 express.json() 以支持 GBK/GB2312 编码请求体
//    （插件端可能使用 GBK 编码的 JSON 请求）
app.use(jsonBodyParser());

// 3. URL-encoded 解析 — 支持表单格式请求体（extended: true 允许嵌套对象）
app.use(express.urlencoded({ extended: true }));

// 4. 静态文件服务 — 托管管理后台前端 SPA（HTML/CSS/JS）
app.use(express.static(path.join(__dirname, "public")));

// ============================================================
// API 路由挂载
// ============================================================

// 管理后台接口（登录、仪表盘、卡密 CRUD、设备列表）— 需 JWT 认证
app.use("/api/admin", adminRouter);
// 公开激活接口 — 卡密验证 + 设备绑定
app.use("/api/activate", activateRouter);
// 公开心跳接口 — 设备在线状态上报
app.use("/api/heartbeat", heartbeatRouter);
// 公开设备管理接口 — 用户自助查看/解绑设备
app.use("/api/devices", devicesRouter);

// ============================================================
// SPA 回退路由
// ============================================================
// 所有非 API 路径（不匹配 /api/*）均返回 index.html，
// 让前端 Vue/原生 JS 路由接管页面渲染。
// 注意：必须放在 API 路由之后，否则会拦截 API 请求。

app.get("*", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ============================================================
// 服务器启动
// ============================================================

function start() {
  // 提前初始化数据库，若数据库文件损坏或无法创建则立即报错退出
  getDb();
  console.log(`[激活后台] 数据库已连接: ${config.dbPath}`);

  app.listen(config.port, () => {
    console.log(`[激活后台] 管理后台已启动: http://localhost:${config.port}`);
    console.log(`[激活后台] 首次使用请运行: npm run setup`);
  });
}

start();
