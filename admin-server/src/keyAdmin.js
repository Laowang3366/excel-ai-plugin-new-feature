/**
 * 卡密管理辅助工具
 *
 * 包含：
 * - normalizeKeyIds: 清洗批量操作中的 ID 数组（去重、类型校验、过滤非法值）
 * - buildExportFilter: 根据导出筛选条件生成 WHERE 子句和标签
 */

/**
 * 标准化卡密 ID 数组
 *
 * 批量接口接收的 ids 数组可能包含：
 * - 字符串数字（如 "123"）
 * - 浮点数（如 123.456）
 * - 重复值
 * - 非法值（0、负数、NaN）
 *
 * 此函数做统一清洗：转为整数、去重、过滤非法值。
 *
 * @param {Array} ids - 原始 ID 数组
 * @returns {number[]} 清洗后的正整数 ID 数组
 */
export function normalizeKeyIds(ids) {
  if (!Array.isArray(ids)) return [];

  return [...new Set(
    ids
      .map((id) => Math.trunc(Number(id)))
      .filter((id) => Number.isInteger(id) && id > 0)
  )];
}

/**
 * 构建导出筛选条件
 *
 * @param {string} filter - 筛选标识（"unused" | "active" 等）
 * @returns {{ label: string, where: string }} label 用于文件名，where 用于 SQL 查询
 */
export function buildExportFilter(filter) {
  if (filter === "unused") {
    return {
      label: "unused",
      where: "WHERE status = 'active' AND used_count = 0",
    };
  }

  // 默认为有效卡密
  return {
    label: "active",
    where: "WHERE status = 'active'",
  };
}
