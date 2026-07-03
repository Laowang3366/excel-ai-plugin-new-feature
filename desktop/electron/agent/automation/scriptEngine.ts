/**
 * 脚本引擎选择与智能执行
 *
 * 组合 Python、JScript、PowerShell 执行能力，为 Excel/WPS 自动化选择可用脚本引擎。
 */

import { executePowerShell } from "./powershell";
import { executeJScript } from "./jscript";
import { executePythonScript, getEmbeddedPythonPath } from "./python";
import { DEFAULT_PROCESS_MAX_BUFFER } from "./processLimits";
import { decodeProcessOutput } from "./stdioEncoding";

/** 脚本引擎类型 */
export type ScriptEngine = "python" | "jscript" | "powershell";

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

  try {
    const { execFileSync } = require("child_process");
    const result = execFileSync(
      "cscript.exe",
      ["//NoLogo", "//E:JScript", "-e", "WScript.Echo('OK')"],
      {
        timeout: 5000,
        maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
        encoding: "buffer",
        stdio: "pipe",
        windowsHide: true,
      }
    );
    if (decodeProcessOutput(result).includes("OK")) {
      _cachedEngine = "jscript";
      return "jscript";
    }
  } catch { /* cscript 不可用 */ }

  _cachedEngine = "powershell";
  return "powershell";
}

/**
 * 智能执行脚本：根据可用引擎自动选择。
 */
export async function executeSmart(
  pythonScript: string,
  jscriptScript: string,
  powershellScript: string,
  timeout = 90000,
  options: { preferPython?: boolean } = {}
): Promise<{ result: string; engine: ScriptEngine }> {
  const preferPython = options.preferPython ?? true;
  const engine = await detectScriptEngine();

  if (preferPython && engine === "python") {
    try {
      const result = await executePythonScript(pythonScript, timeout);
      if (!result || result.trim() === "") {
        throw new Error("Python returned empty output");
      }
      return { result, engine: "python" };
    } catch {
      _cachedEngine = null;
    }
  }

  if (engine === "python" || engine === "jscript") {
    try {
      const result = await executeJScript(jscriptScript, timeout);
      if (!result || result.trim() === "") {
        throw new Error("JScript returned empty output");
      }
      return { result, engine: "jscript" };
    } catch {
      _cachedEngine = null;
    }
  }

  const result = await executePowerShell(powershellScript, timeout);
  return { result, engine: "powershell" };
}
