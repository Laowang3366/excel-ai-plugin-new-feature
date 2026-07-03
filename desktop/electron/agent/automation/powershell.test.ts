import { describe, expect, it } from "vitest";
import { executePowerShell, wrapPowerShellScript } from "./powershell";

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
});
