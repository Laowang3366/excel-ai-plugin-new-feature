/**
 * Office 对象定位解析
 *
 * 关联模块：
 * - types.ts: 返回 OfficeLocator。
 * - officeActionAdapter.ts: 使用解析结果做路由和验证。
 */

import type { OfficeLocator } from "./types";

const INDEXED_KINDS = new Set(["slide", "table", "chart"]);

export function parseOfficeLocator(locator: string): OfficeLocator {
  const separatorIndex = locator.indexOf(":");
  if (separatorIndex < 0) {
    return { kind: locator, value: "" };
  }

  const kind = locator.slice(0, separatorIndex);
  const value = locator.slice(separatorIndex + 1);

  if (kind === "range") {
    const bangIndex = value.indexOf("!");
    if (bangIndex >= 0) {
      return {
        kind,
        value,
        sheetName: value.slice(0, bangIndex),
        address: value.slice(bangIndex + 1),
      };
    }
  }

  if (INDEXED_KINDS.has(kind)) {
    const index = Number(value);
    return Number.isFinite(index) ? { kind, value, index } : { kind, value };
  }

  return { kind, value };
}
