import { decodeProcessOutput } from "./stdioEncoding";
import { DEFAULT_PROCESS_MAX_BUFFER } from "./processLimits";

/**
 * Python 自动化基础能力
 *
 * 负责定位内置/系统 Python，并执行注入 xlwings 上下文的 Python 脚本。
 */

/**
 * 获取内置嵌入式 Python 路径。
 */
export function getEmbeddedPythonPath(): string | null {
  const fs = require("fs");

  for (const candidate of getEmbeddedPythonPathCandidates()) {
    if (fs.existsSync(candidate)) return candidate;
  }

  return null;
}

export function getEmbeddedPythonPathCandidates(): string[] {
  const path = require("path");
  const candidates = [
    path.join(process.resourcesPath || "", "python", "python.exe"),
    path.join(process.cwd(), "python", "python.exe"),
    path.join(__dirname, "..", "python", "python.exe"),
    path.join(__dirname, "..", "..", "python", "python.exe"),
    path.join(__dirname, "..", "..", "..", "python", "python.exe"),
  ];

  return Array.from(new Set(candidates.map((candidate) => path.resolve(candidate))));
}

/**
 * 获取 Python 可执行文件路径（内置优先，系统次之）。
 */
export function getPythonPath(): string {
  const embedded = getEmbeddedPythonPath();
  if (embedded) return embedded;
  return process.platform === "win32" ? "python" : "python3";
}

/**
 * 执行 Python 脚本，返回 stdout。
 */
export function executePythonScript(script: string, timeout = 90000): Promise<string> {
  const { execFile } = require("child_process");
  const pythonPath = getPythonPath();

  const wrappedCode =
    "import json, sys, base64\n" +
    "import xlwings as xw\n" +
    "try:\n" +
    "    app = xw.apps.active\n" +
    "    wb = app.books.active\n" +
    "    ws = wb.sheets.active\n" +
    script + "\n" +
    "except Exception as e:\n" +
    "    print(json.dumps({'error': str(e)}, ensure_ascii=False))\n" +
    "    sys.exit(1)\n";

  const codeB64 = Buffer.from(wrappedCode, "utf-8").toString("base64");

  return new Promise((resolve, reject) => {
    execFile(
      pythonPath,
      ["-c", "import base64,sys;exec(base64.b64decode(sys.argv[1]).decode('utf-8'))", codeB64],
      {
        timeout,
        maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
        encoding: "buffer",
        env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
        windowsHide: true,
      },
      (err: any, stdout: Buffer, stderr: Buffer) => {
        const stdoutText = decodeProcessOutput(stdout).trim();
        const stderrText = decodeProcessOutput(stderr).trim();
        if (err) {
          reject(new Error(stderrText || err.message));
        } else {
          resolve(stdoutText);
        }
      }
    );
  });
}

export interface PlainPythonScriptResult {
  stdout: string;
  stderr: string;
  exitCode: number;
  pythonPath: string;
}

/**
 * 执行通用 Python 脚本。
 *
 * 与 executePythonScript 不同，本函数不注入 Excel/xlwings 上下文，
 * 适合文件级处理（如 docx/openpyxl/pandas 脚本），并通过临时 .py 文件避免 shell 引号转义问题。
 */
export async function executePlainPythonScript(
  script: string,
  timeout = 90000,
  workdir?: string
): Promise<PlainPythonScriptResult> {
  const { execFile } = require("child_process");
  const fs = require("fs");
  const os = require("os");
  const path = require("path");
  const pythonPath = getPythonPath();
  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "excel-ai-python-"));
  const scriptPath = path.join(tempDir, "script.py");

  try {
    await fs.promises.writeFile(scriptPath, script, "utf8");
    return await new Promise((resolve) => {
      execFile(
        pythonPath,
        [scriptPath],
        {
          cwd: workdir,
          timeout,
          maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
          encoding: "buffer",
          env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
          windowsHide: true,
        },
        (err: any, stdout: Buffer, stderr: Buffer) => {
          const code = typeof err?.code === "number" ? err.code : (err ? 1 : 0);
          const stdoutText = decodeProcessOutput(stdout).trim();
          const stderrText = decodeProcessOutput(stderr).trim();
          resolve({
            stdout: stdoutText,
            stderr: stderrText || (err?.message ?? ""),
            exitCode: code,
            pythonPath,
          });
        }
      );
    });
  } finally {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
  }
}

/**
 * 将字符串值以 Python 变量赋值的方式注入脚本。
 */
export function pyVar(name: string, value: string): string {
  const b64 = Buffer.from(value, "utf-8").toString("base64");
  return `${name} = base64.b64decode("${b64}").decode("utf-8")`;
}
