/**
 * Excel/WPS 区域值规范化。
 *
 * 关联模块：
 * - rangeOperations.ts: 写入区域前调用本模块，确保 COM 写入值统一为二维数组。
 */

/**
 * 将单元格、行或矩阵值统一规范化为二维数组。
 */
export function normalize2D(values: unknown): unknown[][] {
  if (values === null || values === undefined) return [[]];
  if (Array.isArray(values) && values.length > 0 && Array.isArray(values[0])) {
    return values;
  }
  if (Array.isArray(values)) {
    return [values];
  }
  return [[values]];
}
