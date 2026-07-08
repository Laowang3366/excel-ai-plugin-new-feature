/**
 * 激活管理后台 — 配置
 *
 * 可通过环境变量覆盖默认值。
 */

const config = {
  /** 服务器端口 */
  port: parseInt(process.env.PORT || "3456", 10),

  /** JWT 密钥（生产环境务必通过环境变量设置） */
  jwtSecret: process.env.JWT_SECRET || "excel-ai-activation-secret-change-in-production",

  /** JWT 有效期（秒） */
  jwtExpiresIn: parseInt(process.env.JWT_EXPIRES_IN || "86400", 10), // 24h

  /** 数据库文件路径（相对于项目根目录） */
  dbPath: process.env.DB_PATH || "./data/activation.db",

  /** 心跳过期时间（秒），超过此时长未收到心跳则标记离线 */
  heartbeatTimeout: parseInt(process.env.HEARTBEAT_TIMEOUT || "120", 10), // 2min

  /** 心跳上报间隔（秒），用于服务端校验 */
  heartbeatInterval: parseInt(process.env.HEARTBEAT_INTERVAL || "300", 10), // 5min

  /** 离线容忍时间（秒），超过此时长未联网验证则锁定 */
  offlineTolerance: parseInt(process.env.OFFLINE_TOLERANCE || "259200", 10), // 72h
};

export default config;
