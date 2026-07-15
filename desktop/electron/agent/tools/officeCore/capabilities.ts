/**
 * Office action 能力声明
 *
 * 关联模块：
 * - officeActionAdapter.ts: 根据能力决定 Open XML 路由或 needsCom。
 * - prompts/templates/scenarios/office-tools.zh-CN.md: 提示词描述这些能力。
 */

import type { OfficeActionApp } from "./types";
import type { OfficeCapability } from "./capabilitiesTypes";
import { EXCEL_CAPABILITIES } from "./excelCapabilities";
import { WORD_CAPABILITIES } from "./wordCapabilities";
import { PRESENTATION_CAPABILITIES } from "./presentationCapabilities";

export type { OfficeCapability } from "./capabilitiesTypes";

export const OFFICE_CAPABILITIES: OfficeCapability[] = [
  ...EXCEL_CAPABILITIES,
  ...WORD_CAPABILITIES,
  ...PRESENTATION_CAPABILITIES,
];

export function findOfficeCapability(
  app: OfficeActionApp,
  operation: string,
): OfficeCapability | undefined {
  return OFFICE_CAPABILITIES.find((item) => item.app === app && item.operation === operation);
}
