/**
 * Python 工具执行器
 *
 * 关联模块：
 * - ../../automation/python: 提供临时脚本文件执行，避免 shell 引号转义问题。
 * - ../../security/sandbox: 复用工作目录白名单检查。
 * - ../registry/python: 模型可见的 python.execute schema。
 */

import * as os from "os";
import type { ToolExecutor } from "../../shared/types";
import { executePlainPythonScript } from "../../automation/python";
import { evaluateCommand } from "../../security/sandbox";
import { validateArgs } from "./validation";

export function addPythonExecutors(target: Map<string, ToolExecutor>): void {
  const executor: ToolExecutor = {
    name: "python.execute",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { code: "string" });
      if (err) return { success: false, error: err };

      const code = args.code as string;
      const requestedWorkdir = (args.workdir as string) || os.homedir();
      const timeout = typeof args.timeout_ms === "number" ? args.timeout_ms : 90000;
      const evaluation = await evaluateCommand("python script.py", requestedWorkdir);

      const result = await executePlainPythonScript(
        code,
        timeout,
        evaluation.cwd.effectiveWorkdir
      );

      return {
        success: result.exitCode === 0,
        data: {
          ...result,
          decision: evaluation.decision,
          workdirRequested: requestedWorkdir,
          workdirEffective: evaluation.cwd.effectiveWorkdir,
          workdirRedirected: evaluation.cwd.redirected,
        },
        error: result.exitCode === 0 ? undefined : (result.stderr || "Python 脚本执行失败"),
      };
    },
  };
  target.set("python.execute", executor);
  // 兼容模型偶发使用的下划线工具名；对外提示词仍只推荐 python.execute。
  target.set("python_execute", executor);
}
