import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { executePowerShell, psVar, wrapPowerShellScript } from "./powershell";

describe("PowerShell automation", () => {
  it("wraps scripts with UTF-8 output setup", () => {
    const wrapped = wrapPowerShellScript("Write-Output 'ok'");

    expect(wrapped.charCodeAt(0)).toBe(0xfeff);
    expect(wrapped).toContain("[Console]::OutputEncoding");
    expect(wrapped).toContain("$OutputEncoding");
    expect(wrapped).toContain("Write-Output 'ok'");
  });

  it("executes long scripts without passing the script through a command-line argument", async () => {
    const script = `$text = '${"x".repeat(40000)}'\nWrite-Output $text.Length`;

    await expect(executePowerShell(script, 10000)).resolves.toBe("40000");
  });

  it.skipIf(process.platform !== "win32")("terminates a registered managed process after timeout", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "excel-ai-ps-timeout-"));
    const observedProcessIdPath = path.join(tempDir, "observed.pid");
    let processId = 0;

    try {
      const script = `
${psVar("_observedProcessIdPath", observedProcessIdPath)}
$managed = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 30') -WindowStyle Hidden -PassThru
[IO.File]::WriteAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE, [string]$managed.Id)
[IO.File]::WriteAllText($_observedProcessIdPath, [string]$managed.Id)
Start-Sleep -Seconds 30
`;
      await expect(executePowerShell(script, 1000)).rejects.toThrow();
      processId = Number.parseInt(await readFile(observedProcessIdPath, "utf8"), 10);
      expect(Number.isSafeInteger(processId)).toBe(true);
      expect(isProcessRunning(processId)).toBe(false);
    } finally {
      if (processId > 0 && isProcessRunning(processId)) {
        try { process.kill(processId, "SIGKILL"); } catch { /* already exited */ }
      }
      await rm(tempDir, { recursive: true, force: true });
    }
  }, 10_000);
});

function isProcessRunning(processId: number): boolean {
  try {
    process.kill(processId, 0);
    return true;
  } catch {
    return false;
  }
}
