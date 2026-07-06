/**
 * Shell 工具执行器
 *
 * 关联 security/sandbox，负责命令策略评估后的实际执行和 shell.execute 工具注册。
 */

import * as os from "os";
import type { ToolExecutionContext, ToolExecutor } from "../../shared/types";
import {
  evaluateCommand,
  runShellSpawn,
  killProcessTree,
  type ShellCommandResult as SandboxShellCommandResult,
  type CommandEvaluation,
} from "../../security/sandbox";
import { validateArgs } from "./validation";
import {
  DEFAULT_SHELL_TIMEOUT_MS,
  MS_PER_SECOND,
  SHELL_WATCHDOG_GRACE_MS,
} from "./shellExecutionLimits";

/** Shell 命令执行结果（复用 sandbox 导出类型，避免双重定义） */
export type ShellCommandResult = SandboxShellCommandResult;

function isCommandEvaluation(value: unknown): value is CommandEvaluation {
  return value !== null &&
    typeof value === "object" &&
    "decision" in value &&
    "evaluation" in value &&
    "cwd" in value;
}

/**
 * 在用户默认 shell 中执行命令。
 */
export async function executeShellCommand(
  evaluation: CommandEvaluation,
  command: string,
  workdir: string,
  timeoutMs: number
): Promise<ShellCommandResult> {
  if (evaluation.decision === "forbidden") {
    return {
      stdout: "",
      stderr: evaluation.violationMessage || "命令被安全策略拒绝",
      exitCode: 126,
    };
  }

  return new Promise((resolve) => {
    let settled = false;
    const finish = (r: ShellCommandResult) => {
      if (settled) return;
      settled = true;
      clearTimeout(watchdog);
      resolve(r);
    };

    const child = runShellSpawn(command, workdir, timeoutMs, (r) => finish(r));
    child.on("error", () => finish({ stdout: "", stderr: "无法启动子进程", exitCode: 1 }));

    const watchdog = setTimeout(() => {
      killProcessTree(child).finally(() => finish({
        stdout: "",
        stderr: `命令执行超时（${timeoutMs / MS_PER_SECOND}s），已强杀进程树`,
        exitCode: -1,
      }));
    }, timeoutMs + SHELL_WATCHDOG_GRACE_MS);
  });
}

export function addShellExecutors(target: Map<string, ToolExecutor>): void {
  target.set("shell.execute", {
    name: "shell.execute",
    execute: async (args: Record<string, unknown>, context?: ToolExecutionContext) => {
      const err = validateArgs(args, { command: "string" });
      if (err) return { success: false, error: err };

      const command = args.command as string;
      const requestedWorkdir = (args.workdir as string) || os.homedir();
      const timeout = typeof args.timeout_ms === "number" ? args.timeout_ms : DEFAULT_SHELL_TIMEOUT_MS;
      const evaluation = isCommandEvaluation(context?.sandboxEvaluation)
        ? context.sandboxEvaluation
        : await evaluateCommand(command, requestedWorkdir);

      if (evaluation.decision === "forbidden") {
        return {
          success: false,
          error: evaluation.violationMessage || "命令被安全策略拒绝",
          data: {
            decision: "forbidden",
            violations: evaluation.evaluation.violations.map((h) => ({
              matched: h.matchedPrefix,
              justification: h.rule.justification,
            })),
            requestedWorkdir,
            effectiveWorkdir: evaluation.cwd.effectiveWorkdir,
            workdirRedirected: evaluation.cwd.redirected,
          },
        };
      }

      try {
        const stat = await import("fs").then((fs) =>
          fs.promises.stat(evaluation.cwd.effectiveWorkdir)
        );
        if (!stat.isDirectory()) {
          return { success: false, error: `工作目录不是有效目录: ${evaluation.cwd.effectiveWorkdir}` };
        }
      } catch {
        return { success: false, error: `工作目录不存在: ${evaluation.cwd.effectiveWorkdir}` };
      }

      const result = await executeShellCommand(
        evaluation,
        command,
        evaluation.cwd.effectiveWorkdir,
        timeout
      );
      return {
        success: result.exitCode === 0,
        data: {
          ...result,
          decision: evaluation.decision,
          workdirRequested: requestedWorkdir,
          workdirEffective: evaluation.cwd.effectiveWorkdir,
          workdirRedirected: evaluation.cwd.redirected,
          matchedRules: evaluation.evaluation.hits.map((h) => ({
            matched: h.matchedPrefix,
            decision: h.rule.decision,
            justification: h.rule.justification,
          })),
        },
      };
    },
  });
}
