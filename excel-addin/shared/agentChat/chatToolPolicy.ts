import { TOOL_DEFINITIONS, TOOL_DEFINITION_MAP } from "../tools";
import type { ToolDefinition, ToolName } from "../tools/types";
import {
  dispositionForRisk,
  type ApprovalDisposition,
} from "./approvalPolicy";

/**
 * Fresh shallow copies of all registry tools for this chat turn.
 * Does not mutate TOOL_DEFINITIONS / TOOL_DEFINITION_MAP.
 */
export function listChatTools(): ToolDefinition[] {
  return TOOL_DEFINITIONS.map((tool) => ({
    ...tool,
    parameters: tool.parameters,
  }));
}

export function classifyChatTool(name: string): {
  disposition: ApprovalDisposition;
  definition?: ToolDefinition;
} {
  if (!Object.prototype.hasOwnProperty.call(TOOL_DEFINITION_MAP, name)) {
    return { disposition: "deny" };
  }
  const definition = TOOL_DEFINITION_MAP[name as ToolName];
  return {
    disposition: dispositionForRisk(definition?.riskLevel),
    definition,
  };
}
