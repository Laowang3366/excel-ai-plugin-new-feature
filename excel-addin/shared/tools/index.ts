import { TOOL_DEFINITIONS } from "./definitions";
import type { ToolDefinition } from "./types";

export { TOOL_DEFINITIONS } from "./definitions";
export const TOOL_DEFINITION_MAP = Object.fromEntries(
  TOOL_DEFINITIONS.map((tool) => [tool.name, tool]),
) as Record<ToolDefinition["name"], ToolDefinition>;
export { ToolExecutor } from "./executor";
export type {
  CellValue,
  FormulaWriteArgs,
  RangeWriteArgs,
  RiskLevel,
  ToolCall,
  ToolDefinition,
  ToolFailure,
  ToolName,
  ToolResult,
  ToolSuccess,
} from "./types";
