/**
 * 脚本引擎选择与智能执行
 *
 * 在 Python 与 PowerShell 之间选择 Excel/WPS 自动化引擎。
 */

import { executePowerShell } from "./powershell";
import { executePythonScript, getEmbeddedPythonPath } from "./python";
import { DEFAULT_PROCESS_MAX_BUFFER } from "./processLimits";

/** 脚本引擎类型 */
export type ScriptEngine = "python" | "powershell";

/** 缓存可用的脚本引擎，避免重复检测 */
export let _cachedEngine: ScriptEngine | null = null;

/** 重置引擎缓存（用于 fallback 后重新检测） */
export function resetEngineCache(): void {
  _cachedEngine = null;
}

/**
 * 检测最佳可用脚本引擎。
 */
export async function detectScriptEngine(): Promise<ScriptEngine> {
  if (_cachedEngine) return _cachedEngine;

  const embeddedPython = getEmbeddedPythonPath();
  if (embeddedPython) {
    try {
      const { execFileSync } = require("child_process");
      execFileSync(embeddedPython, ["-c", "import xlwings; print('OK')"], {
        timeout: 5000,
        maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
        encoding: "buffer",
        stdio: "pipe",
        windowsHide: true,
      });
      _cachedEngine = "python";
      return "python";
    } catch { /* 内置 Python 不可用，继续 */ }
  }

  try {
    const { execFileSync } = require("child_process");
    const pyCmd = process.platform === "win32" ? "python" : "python3";
    execFileSync(pyCmd, ["-c", "import xlwings; print('OK')"], {
      timeout: 5000,
      maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
      encoding: "buffer",
      stdio: "pipe",
      windowsHide: true,
    });
    _cachedEngine = "python";
    return "python";
  } catch { /* 系统 Python 不可用，继续 */ }

  _cachedEngine = "powershell";
  return "powershell";
}

/**
 * 优先执行 Python/xlwings，失败时回退到 PowerShell COM。
 */
export async function executeSmart(
  pythonScript: string,
  powershellScript: string,
  timeout = 90000,
): Promise<{ result: string; engine: ScriptEngine }> {
  const engine = await detectScriptEngine();

  if (engine === "python") {
    try {
      const result = await executePythonScript(pythonScript, timeout);
      if (!result || result.trim() === "") {
        throw new Error("Python returned empty output");
      }
      return { result, engine: "python" };
    } catch {
      _cachedEngine = "powershell";
    }
  }

  const result = await executePowerShell(powershellScript, timeout);
  return { result, engine: "powershell" };
}
