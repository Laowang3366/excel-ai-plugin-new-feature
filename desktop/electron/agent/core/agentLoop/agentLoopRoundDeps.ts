import type { ToolExecutor } from "../../shared/types";
import { runAgentLoopRounds, type AgentLoopRunnerInput } from "./agentLoopRunner";
import type { ToolApprovalConfig } from "./toolExecutor";

type AgentLoopWithDepsInput = Omit<AgentLoopRunnerInput, "toolExecutors" | "approvalConfig"> & {
  toolExecutors?: Map<string, ToolExecutor>;
  permissionMode?: ToolApprovalConfig["permissionMode"];
  requestToolApproval?: ToolApprovalConfig["requestToolApproval"];
};

export async function runAgentLoopWithDeps(input: AgentLoopWithDepsInput): Promise<void> {
  await runAgentLoopRounds({
    ...input,
    toolExecutors: input.toolExecutors!,
    approvalConfig: {
      permissionMode: input.permissionMode || "normal",
      requestToolApproval: input.requestToolApproval,
    },
  });
}
