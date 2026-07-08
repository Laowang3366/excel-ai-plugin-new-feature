/**
 * JWT 认证中间件
 *
 * 验证请求头中的 Bearer token，解码后将管理员信息注入 req.admin。
 */

import jwt from "jsonwebtoken";
import config from "../config.js";

/**
 * 生成 JWT token
 * @param {object} payload - 要签入 token 的数据
 * @returns {string} JWT token
 */
export function signToken(payload) {
  return jwt.sign(payload, config.jwtSecret, { expiresIn: config.jwtExpiresIn });
}

/**
 * 验证 JWT token
 * @param {string} token
 * @returns {object|null} 解码后的 payload，无效返回 null
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
 * 验证通过后将 { id, username } 注入 req.admin。
 */
export function requireAdmin(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "未提供认证令牌" });
  }

  const token = authHeader.slice(7);
  const decoded = verifyToken(token);
  if (!decoded) {
    return res.status(401).json({ error: "认证令牌无效或已过期" });
  }

  req.admin = { id: decoded.id, username: decoded.username };
  next();
}
