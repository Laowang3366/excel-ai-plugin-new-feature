/**
 * 沙箱入口 — shell.execute 的统一前置评估 + spawn 包装
 *
 * 调用流：
 *   executors.executeShellCommand
 *     → sandbox.evaluateCommand(command, workdir)
 *        ├─ parseCommand：切子命令 + token
 *        ├─ ExecPolicy.evaluate：allow / prompt / forbidden
 *        └─ checkWorkdir：workdir 白名单 + 重定向
 *     → 拿结果走审批或拒绝；通过后调 runShellSpawn 执行
 *
 * 与 Codex `core/src/tools/sandboxing.rs` 的 `Sandboxable` 入口对齐：
 *   evaluate → decision → choose path (deny / approval / spawn)
 */

import * as os from "os";
import { type ChildProcess, execFile } from "child_process";
import { decodeProcessOutput } from "../../automation/stdioEncoding";
import { parseCommand, type ParsedCommand } from "./parseCommand";
import {
  ExecPolicy,
  checkWorkdir,
  defaultWritableRoots,
  type ExecPolicyEvaluation,
  type PrefixRule,
  type Decision,
} from "./execPolicy";
import { DEFAULT_RULES } from "./defaultRules";
import { audit } from "./audit";

export {
  parseCommand,
  ExecPolicy,
  checkWorkdir,
  defaultWritableRoots,
  DEFAULT_RULES,
};
export type { ParsedCommand, ExecPolicyEvaluation, PrefixRule, Decision };

export const SHELL_STDOUT_MAX_CHARS = 50000;
export const SHELL_STDERR_MAX_CHARS = 10000;

// ============================================================
// 单例策略引擎：默认规则 + 用户自定义规则（来自 settingsManager）
// ============================================================

let extraWritableRoots: string[] = [];
const engine = new ExecPolicy(DEFAULT_RULES);

/** 暴露给设置模块：热更新用户规则 */
export function setUserRules(rules: PrefixRule[]): void {
  engine.setRules([...rules, ...DEFAULT_RULES]);
}

/** 暴露给设置模块：热更新额外可写根（cwd 白名单展开项） */
export function setExtraWritableRoots(roots: string[]): void {
  extraWritableRoots = roots;
}

// ============================================================
// 评估结果
// ============================================================

export interface CommandEvaluation {
  /** 策略最终决策 */
  decision: Decision;
  /** 命中信息 */
  evaluation: ExecPolicyEvaluation;
  /** workdir 检查结果 */
  cwd: { allowed: boolean; effectiveWorkdir: string; redirected: boolean };
  /** 已切片的子命令（供执行用） */
  parsed: ParsedCommand[];
  /** forbidden 的理由文本（用于模型回包与审批对话框） */
  violationMessage?: string;
}

/**
 * 评估一条 shell.execute 命令
 *
 * 不执行、不审批——只产出 evaluation。审批/拒绝由 toolExecutor 决定。
 */
export async function evaluateCommand(
  command: string,
  workdir: string
): Promise<CommandEvaluation> {
  const parsed = parseCommand(command);
  const evaluation = engine.evaluate(parsed);
  const roots = [...defaultWritableRoots(), ...extraWritableRoots];
  const cwd = checkWorkdir(workdir, roots, os.tmpdir());

  // 审计：策略命中即记一条 decision；forbidden 单独记一条 violation
  if (evaluation.hits.length > 0 || evaluation.unparseable.length > 0) {
    await audit({
      type: "decision",
      ts: new Date().toISOString(),
      command,
      decision: evaluation.decision,
      hits: evaluation.hits.map((h) => ({
        matchedPrefix: h.matchedPrefix,
        decision: h.rule.decision,
        justification: h.rule.justification,
      })),
      requestedWorkdir: workdir,
      effectiveWorkdir: cwd.effectiveWorkdir,
      redirected: cwd.redirected,
    });
  }
  if (evaluation.violations.length > 0) {
    await audit({
      type: "violation",
      ts: new Date().toISOString(),
      command,
      decision: "forbidden",
      hits: evaluation.violations.map((h) => ({
        matchedPrefix: h.matchedPrefix,
        decision: h.rule.decision,
        justification: h.rule.justification,
      })),
    });
  }
  if (cwd.redirected) {
    await audit({
      type: "workdir_redirect",
      ts: new Date().toISOString(),
      command,
      requestedWorkdir: workdir,
      effectiveWorkdir: cwd.effectiveWorkdir,
      redirected: true,
    });
  }

  let violationMessage: string | undefined;
  if (evaluation.violations.length > 0) {
    const reasons = evaluation.violations
      .map((h) => h.rule.justification || h.matchedPrefix.join(" "))
      .filter(Boolean);
    violationMessage = `命令被安全策略拒绝：${reasons.join("；") || "匹配禁止规则"}`;
  }

  return {
    decision: evaluation.decision,
    evaluation,
    cwd,
    parsed,
    violationMessage,
  };
}

// ============================================================
// spawn 包装：Env allowlist + EncodedCommand 标准化 + 进程树强杀
// ============================================================

/** Shell 执行结果（与 executors.ts 原有结构一致） */
export interface ShellCommandResult {
  stdout: string;
  stderr: string;
  exitCode: number;
}

/** 危险环境变量白名单 —— 对应 Codex `WINDOWS_SANDBOX_WRAPPER_SETUP_ENV_ALLOWLIST` */
const ENV_ALLOWLIST = new Set(
  (process.platform === "win32"
    ? ["USERNAME", "USERPROFILE", "SystemRoot", "TEMP", "TMP", "PATH", "PATHEXT"]
    : ["HOME", "PATH", "SHELL", "LANG", "LC_ALL", "LC_CTYPE"])
    .map((k) => k.toUpperCase())
);

/** 清洗子进程环境（仅保留 ENV_ALLOWLIST）——用于 shell.execute 路径 */
function sanitizeEnv(env: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const clean: NodeJS.ProcessEnv = {};
  for (const [k, v] of Object.entries(env)) {
    if (ENV_ALLOWLIST.has(k.toUpperCase())) clean[k] = v;
  }
  return clean;
}

/**
 * 执行 shell 命令
 *
 * 与原 executeShellCommand 签名一致，但：
 * - Windows 一律改走 `-EncodedCommand`（Base64 UTF-16LE），避免 `-Command` 明文拼接注入
 * - 清洗环境变量（仅保留白名单）—— 对应 Codex process-hardening 的 env 剥离
 * - Linux/macOS 用 `/bin/bash -lc command`（沿用原行为）
 * - 通过 ChildProcess 持有句柄；调用方可在 timeout 时调 killProcessTree 强杀
 *
 * 注意：Excel COM 工具不走这里（它们走 excelBridge 的 executePowerShell，
 * 保留完整 PSModulePath 等环境以连接 Excel）。
 */
export function runShellSpawn(
  command: string,
  workdir: string,
  timeoutMs: number,
  onDone: (result: ShellCommandResult) => void
): ChildProcess {
  const isWin = process.platform === "win32";
  const shell = isWin ? "powershell.exe" : "/bin/bash";
  let shellArgs: string[];
  if (isWin) {
    // 在脚本首行强制 UTF-8 输出编码，避免中文乱码（与 automation/powershell 保持一致）
    const full = `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n$OutputEncoding = [System.Text.Encoding]::UTF8\n${command}`;
    const encoded = Buffer.from(full, "utf16le").toString("base64");
    shellArgs = ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded];
  } else {
    shellArgs = ["-lc", command];
  }

  return execFile(shell, shellArgs, {
    cwd: workdir,
    timeout: timeoutMs,
    maxBuffer: 1024 * 1024,
    encoding: "buffer",
    windowsHide: true,
    env: {
      ...sanitizeEnv(process.env),
      PYTHONIOENCODING: "utf-8",
      PYTHONUTF8: "1",
    },
  }, (error, stdout, stderr) => {
    const stdoutText = decodeProcessOutput(stdout);
    const stderrText = decodeProcessOutput(stderr);
    void audit({
      type: error ? "execute_failure" : "execute",
      ts: new Date().toISOString(),
      command,
      exitCode: error ? -1 : 0,
      error: error?.message,
    });

    if (error) {
      const exitCode = (error as NodeJS.ErrnoException).code === "ETIMEDOUT"
        ? -1
        : ((error as NodeJS.ErrnoException).code as unknown as number) ?? 1;
      onDone({
        stdout: stdoutText.slice(0, SHELL_STDOUT_MAX_CHARS),
        stderr: stderrText.slice(0, SHELL_STDERR_MAX_CHARS) || error.message,
        exitCode: typeof exitCode === "number" ? exitCode : 1,
      });
    } else {
      onDone({
        stdout: stdoutText.slice(0, SHELL_STDOUT_MAX_CHARS),
        stderr: stderrText.slice(0, SHELL_STDERR_MAX_CHARS),
        exitCode: 0,
      });
    }
  });
}

/**
 * 强杀整条进程树（对应 Codex linux-sandbox 的 unshare-pid + 杀树；Windows 用 wmic/taskkill）
 *
 * 在 timeout 触发时调用：先 child.kill('SIGKILL')，再针对 Windows 用 taskkill /T 杀子树。
 */
export async function killProcessTree(child: ChildProcess): Promise<void> {
  if (child.exitCode !== null || child.killed) return;
  try {
    child.kill("SIGKILL");
  } catch {
    /* ignore */
  }
  if (process.platform === "win32" && child.pid) {
    try {
      await new Promise<void>((resolve) => {
        execFile(
          "taskkill",
          ["/T", "/F", "/PID", String(child.pid)],
          { windowsHide: true },
          () => resolve()
        );
      });
    } catch {
      /* ignore */
    }
  }
}
