/**
 * 北京时间（UTC+8）日期时间格式化工具
 *
 * SQLite 的 datetime('now') 返回的是 UTC 时间（ISO-8601 格式 YYYY-MM-DD HH:MM:SS），
 * 管理后台的运营人员在中国时区，需要将所有时间戳显示为北京时间。
 *
 * 转换策略：
 * - 解析 UTC 时间字符串为 Date 对象
 * - 手动加上 8 小时偏移（而非依赖 toLocaleString，因为 Node.js 环境可能时区配置不同）
 * - 输出格式统一为 "YYYY-MM-DD HH:MM:SS"
 *
 * 适用范围：
 * - created_at, updated_at, activated_at, last_heartbeat, expires_at
 */

/** 北京时间偏移（毫秒） */
const BEIJING_OFFSET_MS = 8 * 60 * 60 * 1000;

/** 需要转换的字段名列表 */
const TIMESTAMP_FIELDS = ["created_at", "updated_at", "activated_at", "last_heartbeat", "expires_at"];

/**
 * 将单个 UTC 时间戳字符串转换为北京时间格式
 *
 * @param {string|null} value - SQLite UTC 时间字符串 "YYYY-MM-DD HH:MM:SS"
 * @returns {string|null} 北京时间字符串，无效输入原样返回
 */
export function formatBeijingDateTime(value) {
  if (!value) return value;

  // 匹配 SQLite 的时间格式（兼容 "T" 分隔符）
  const match = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/.exec(String(value));
  if (!match) return value;

  const [, year, month, day, hour, minute, second] = match;
  // 将 UTC 时间解析为毫秒时间戳
  const utcMs = Date.UTC(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  );

  // 加上 8 小时偏移并格式化为 ISO 字符串，再截取为 SQLite 格式
  return new Date(utcMs + BEIJING_OFFSET_MS).toISOString().slice(0, 19).replace("T", " ");
}

/**
 * 将记录对象中的所有时间戳字段转换为北京时间
 *
 * @param {object|null} record - 数据库查询结果行
 * @returns {object|null} 转换后的新对象（不修改原对象）
 */
export function withBeijingDateTimes(record) {
  if (!record) return record;

  const converted = { ...record };
  for (const field of TIMESTAMP_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(converted, field)) {
      converted[field] = formatBeijingDateTime(converted[field]);
    }
  }
  return converted;
}
