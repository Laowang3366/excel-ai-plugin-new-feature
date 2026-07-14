/**
 * 工具定义聚合入口
 *
 * 按领域组合工具定义，保持对外 ALL_TOOL_DEFINITIONS 顺序不变。
 */

import type { ToolDefinition } from "../../shared/types";
import { WORKBOOK_TOOL_DEFINITIONS } from "./workbook";
import { RANGE_TOOL_DEFINITIONS } from "./range";
import { FORMULA_TOOL_DEFINITIONS } from "./formula";
import { SHEET_TOOL_DEFINITIONS } from "./sheet";
import { MACRO_TOOL_DEFINITIONS } from "./macro";
import { UI_TOOL_DEFINITIONS } from "./ui";
import { FILE_TOOL_DEFINITIONS } from "./file";
import { KNOWLEDGE_TOOL_DEFINITIONS } from "./knowledge";
import { WEB_TOOL_DEFINITIONS } from "./web";
import { MEMORY_TOOL_DEFINITIONS } from "./memory";
import { OFFICE_TOOL_DEFINITIONS } from "./office";
import { OCR_TOOL_DEFINITIONS } from "./ocr";

/** 所有工具定义 */
export const ALL_TOOL_DEFINITIONS: ToolDefinition[] = [
  WORKBOOK_TOOL_DEFINITIONS[0],
  ...RANGE_TOOL_DEFINITIONS,
  ...FORMULA_TOOL_DEFINITIONS,
  ...MACRO_TOOL_DEFINITIONS,
  ...SHEET_TOOL_DEFINITIONS,
  ...UI_TOOL_DEFINITIONS,
  ...FILE_TOOL_DEFINITIONS,
  ...WORKBOOK_TOOL_DEFINITIONS.slice(1, 4),
  ...WORKBOOK_TOOL_DEFINITIONS.slice(4),
  ...KNOWLEDGE_TOOL_DEFINITIONS,
  ...WEB_TOOL_DEFINITIONS,
  ...OCR_TOOL_DEFINITIONS,
  ...MEMORY_TOOL_DEFINITIONS,
  ...OFFICE_TOOL_DEFINITIONS,
];

/** 工具定义映射 */
export const TOOL_DEFINITIONS_MAP = new Map<string, ToolDefinition>(
  ALL_TOOL_DEFINITIONS.flatMap((tool) => [
    [tool.name, tool] as const,
    ...getToolNameAliases(tool.name).map((alias) => [alias, tool] as const),
  ])
);

export function getToolNameAliases(toolName: string): string[] {
  if (!toolName.includes(".")) return [];
  return Array.from(new Set([
    toolName.replace(/\./g, "_"),
    toolName.replace(/\.(?=[^.]+$)/, "_"),
  ].filter((alias) => alias !== toolName)));
}
