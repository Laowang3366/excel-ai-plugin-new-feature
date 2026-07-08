/**
 * JWT 认证中间件
 *
 * 验证请求头中的 Bearer token，解码后将管理员信息注入 req.admin。
 *
 * 安全设计：
 * - 密码采用 bcrypt 哈希（在 admin.js 登录处），此处只负责 JWT 生命周期。
 * - JWT 负载仅包含 id 和 username，不包含敏感信息。
 * - 每次请求均验证签名和有效期，无状态认证（无需查库验证 token）。
 * - token 泄露风险由有效期控制（默认 24h），生产环境建议缩短并配合 HTTPS。
 */

import jwt from "jsonwebtoken";
import config from "../config.js";

/**
 * 生成 JWT token
 *
 * @param {object} payload - 要签入 token 的数据（通常含 id, username）
 * @returns {string} 签名的 JWT 字符串
 */
export function signToken(payload) {
  return jwt.sign(payload, config.jwtSecret, { expiresIn: config.jwtExpiresIn });
}

/**
 * 验证 JWT token
 *
 * 使用配置的 jwtSecret 校验签名和有效期。
 * 不抛出异常，验证失败统一返回 null 以便上层决策。
 *
 * @param {string} token - 原始 JWT 字符串
 * @returns {object|null} 解码后的 payload（含 iat, exp），无效或过期返回 null
 */
export function verifyToken(token) {
  try {
    return jwt.verify(token, config.jwtSecret);
  } catch {
    return null;
  }
}

/**
 * Express 中间件 — 要求请求携带有效 JWT
 *
 * 在 Authorization 头中以 "Bearer <token>" 格式传递。
 * 验证通过后将 { id, username } 注入 req.admin，供后续路由使用。
 * 验证失败返回 401 JSON，不会调用 next()。
 *
 * @param {import('express').Request} req
 * @param {import('express').Response} res
 * @param {import('express').NextFunction} next
 */
export function requireAdmin(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "未提供认证令牌" });
  }

  // 提取 "Bearer " 后面的 token 部分
  const token = authHeader.slice(7);
  const decoded = verifyToken(token);
  if (!decoded) {
    return res.status(401).json({ error: "认证令牌无效或已过期" });
  }

  // 将管理员身份信息注入请求对象，下游路由可通过 req.admin 访问
  req.admin = { id: decoded.id, username: decoded.username };
  next();
}
