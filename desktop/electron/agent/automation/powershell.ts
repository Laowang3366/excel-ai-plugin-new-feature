import { DEFAULT_PROCESS_MAX_BUFFER } from "./processLimits";
import { decodeProcessOutput } from "./stdioEncoding";

/**
 * PowerShell 自动化基础能力
 *
 * 被 Excel、Word、PowerPoint COM 桥复用，负责执行 PowerShell 和安全注入字符串变量。
 */

/**
 * 执行 PowerShell 命令，返回 stdout。
 */
export async function executePowerShell(script: string, timeout = 90000): Promise<string> {
  const { execFile } = require("child_process");
  const fs = require("fs");
  const os = require("os");
  const path = require("path");
  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "excel-ai-ps-"));
  const scriptPath = path.join(tempDir, "script.ps1");
  const managedProcessIdPath = path.join(tempDir, "managed-process.pid");

  try {
    await fs.promises.writeFile(scriptPath, wrapPowerShellScript(script), "utf8");
    return await new Promise((resolve, reject) => {
      execFile(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
        {
          timeout,
          maxBuffer: DEFAULT_PROCESS_MAX_BUFFER,
          encoding: "buffer",
          windowsHide: true,
          env: { ...process.env, WENGGE_MANAGED_PROCESS_ID_FILE: managedProcessIdPath },
        },
        async (err: any, stdout: Buffer, stderr: Buffer) => {
          const stdoutText = decodeProcessOutput(stdout).trim();
          const stderrText = decodeProcessOutput(stderr).trim();
          if (err) {
            await terminateManagedProcess(fs, managedProcessIdPath);
            reject(new Error(stderrText || err.message));
          } else {
            resolve(stdoutText);
          }
        }
      );
    });
  } finally {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
  }
}

async function terminateManagedProcess(fs: any, processIdPath: string): Promise<void> {
  try {
    const content = String(await fs.promises.readFile(processIdPath, "utf8"));
    const processIds = [...new Set<number>((content.match(/\d+/g) || []).map(Number))]
      .filter((processId) => Number.isSafeInteger(processId) && processId > 0);
    for (const processId of processIds) {
      try { process.kill(processId, "SIGKILL"); } catch { /* already exited */ }
    }
  } catch { return; }
}

export function wrapPowerShellScript(script: string): string {
  return `\ufeff[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n$OutputEncoding = [System.Text.Encoding]::UTF8\n${script}`;
}

/**
 * 将字符串值以 PowerShell 变量赋值的方式注入脚本。
 */
export function psVar(name: string, value: string): string {
  const b64 = Buffer.from(value, "utf16le").toString("base64");
  return `$${name} = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String('${b64}'))`;
}
