/**
 * 激活管理后台 — 集中配置
 *
 * 所有可调参数汇集于此，通过环境变量覆盖默认值以适配不同部署环境。
 * 生产部署时必须设置 JWT_SECRET，其余可按需调整。
 */

const config = {
  /**
   * 服务器监听端口
   * @env PORT
   * @default "3456"
   */
  port: parseInt(process.env.PORT || "3456", 10),

  /**
   * JWT 签名密钥
   * 生产环境务必通过环境变量 JWT_SECRET 设置一个足够长且随机的字符串，
   * 不要使用默认值，否则存在 token 伪造风险。
   * @env JWT_SECRET
   */
  jwtSecret: process.env.JWT_SECRET || "excel-ai-activation-secret-change-in-production",

  /**
   * JWT 有效期（秒）
   * 默认 86400 秒（24 小时）。过期后用户需重新登录。
   * @env JWT_EXPIRES_IN
   */
  jwtExpiresIn: parseInt(process.env.JWT_EXPIRES_IN || "86400", 10),

  /**
   * SQLite 数据库文件路径
   * 支持相对路径（相对于 CWD）和绝对路径。
   * @env DB_PATH
   */
  dbPath: process.env.DB_PATH || "./data/activation.db",

  /**
   * 心跳超时阈值（秒）
   * 设备最后心跳距当前超过此值则判定为离线。
   * 应与客户端上报间隔匹配，通常设为上报间隔的 2~3 倍以容忍网络抖动。
   * @env HEARTBEAT_TIMEOUT
   * @default 120（2 分钟）
   */
  heartbeatTimeout: parseInt(process.env.HEARTBEAT_TIMEOUT || "120", 10),

  /**
   * 心跳上报间隔（秒）
   * 客户端应在此间隔内上报一次心跳，服务端用作 delta 计算的参考上限。
   * @env HEARTBEAT_INTERVAL
   * @default 300（5 分钟）
   */
  heartbeatInterval: parseInt(process.env.HEARTBEAT_INTERVAL || "300", 10),

  /**
   * 离线容忍时间（秒）
   * 客户端超过此时长未联网验证则锁定功能。
   * 仅对客户端有意义，服务端记录该值供下发参考。
   * @env OFFLINE_TOLERANCE
   * @default 259200（72 小时）
   */
  offlineTolerance: parseInt(process.env.OFFLINE_TOLERANCE || "259200", 10),
};

export default config;
