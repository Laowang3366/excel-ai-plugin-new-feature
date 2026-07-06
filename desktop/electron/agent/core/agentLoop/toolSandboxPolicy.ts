import * as os from "os";

import { evaluateCommand, type CommandEvaluation } from "../../security/sandbox";

export interface ToolSandboxPolicyResult {
  evaluation: CommandEvaluation | null;
  justification?: string;
  forcedForbidden: boolean;
  forcedApproval: boolean;
}

export async function evaluateToolSandboxPolicy(
  canonicalToolName: string,
  argumentsJson: string
): Promise<ToolSandboxPolicyResult> {
  const initial: ToolSandboxPolicyResult = {
    evaluation: null,
    justification: undefined,
    forcedForbidden: false,
    forcedApproval: false,
  };
  if (canonicalToolName !== "shell.execute") return initial;

  try {
    const parsedArgs = parseToolArgumentsForSandbox(argumentsJson);
    const cmd = (parsedArgs.command as string) || "";
    const workdir = (parsedArgs.workdir as string) || os.homedir();
    const evaluation = await evaluateCommand(cmd, workdir);

    if (evaluation.decision === "forbidden") {
      return {
        evaluation,
        justification: evaluation.violationMessage,
        forcedForbidden: true,
        forcedApproval: false,
      };
    }

    if (evaluation.decision === "prompt") {
      return {
        evaluation,
        justification: collectPromptJustifications(evaluation),
        forcedForbidden: false,
        forcedApproval: true,
      };
    }

    return { ...initial, evaluation };
  } catch {
    return {
      evaluation: null,
      justification: "命令策略评估异常，需要用户确认",
      forcedForbidden: false,
      forcedApproval: true,
    };
  }
}

function parseToolArgumentsForSandbox(argumentsJson: string): Record<string, unknown> {
  try {
    return JSON.parse(argumentsJson || "{}");
  } catch {
    return { _raw: argumentsJson };
  }
}

function collectPromptJustifications(evaluation: CommandEvaluation): string {
  const hits = evaluation.evaluation.hits ?? [];
  return hits
    .filter((hit) => hit.rule.decision === "prompt")
    .map((hit) => hit.rule.justification || hit.matchedPrefix.join(" "))
    .filter(Boolean)
    .join("；")
    || "命中安全策略 prompt 规则，需要用户确认";
}
