import type { ToolCall, ToolName, ToolResult } from "../tools/types";
import { TOOL_DEFINITION_MAP } from "../tools";
import type { AgentToolExecutor } from "./chatReadOnlyTools";
import { ApprovalGate } from "./approvalGate";
import {
  deniedToolError,
  dispositionForRisk,
  rejectedToolError,
} from "./approvalPolicy";
import { previewFromToolCall } from "./approvalPreview";
import type { RiskLevel } from "../tools/types";

export type ToolCallContext = {
  toolCallId?: string;
  round?: number;
};

/**
 * Per-call approval wrapper. Safe tools execute immediately; moderate/dangerous
 * await ApprovalGate. Raw args stay local and are only passed to inner after approve.
 */
export class ApprovingToolExecutor implements AgentToolExecutor {
  constructor(
    private readonly inner: AgentToolExecutor,
    private readonly gate: ApprovalGate,
    private readonly getToolCallContext?: () => ToolCallContext | undefined,
  ) {}

  async execute(call: ToolCall): Promise<ToolResult> {
    const name = call.name;
    if (!Object.prototype.hasOwnProperty.call(TOOL_DEFINITION_MAP, name)) {
      return deny(name);
    }
    const def = TOOL_DEFINITION_MAP[name as ToolName];
    const risk = def?.riskLevel as RiskLevel | undefined;
    const disposition = dispositionForRisk(risk);
    if (disposition === "deny" || !def || !risk) {
      return deny(name);
    }
    if (disposition === "direct") {
      return this.inner.execute(call);
    }

    // moderate / dangerous — preview only public fields into gate
    const { argsPreview, destructive, impactHint } = previewFromToolCall(call);
    const ctx = this.getToolCallContext?.();
    const decision = await this.gate.request({
      name,
      riskLevel: risk,
      destructive,
      argsPreview,
      impactHint,
      toolCallId: ctx?.toolCallId,
      round: ctx?.round,
    });
    if (decision === "rejected") {
      return {
        ok: false,
        tool: name as ToolName,
        error: rejectedToolError(name),
      };
    }
    // approved — pass original raw call to host executor
    return this.inner.execute(call);
  }
}

function deny(name: string): ToolResult {
  return {
    ok: false,
    tool: name as ToolName,
    error: deniedToolError(name),
  };
}
