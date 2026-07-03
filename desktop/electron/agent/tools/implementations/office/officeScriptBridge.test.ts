import { beforeEach, describe, expect, test, vi } from "vitest";
import { OfficeScriptBridge } from "./officeScriptBridge";
import { executePowerShell } from "../../../automation/powershell";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

describe("OfficeScriptBridge", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("injects script code with psVar and executes against an existing Word app", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      success: true,
      app: "word",
      progId: "Word.Application",
      engine: "powershell",
      output: "ok",
    }));

    const bridge = new OfficeScriptBridge();
    const result = await bridge.executeScript("word", "$app.Name");

    const script = executePowerShellMock.mock.calls[0][0];
    expect(result).toMatchObject({ success: true, app: "word", output: "ok" });
    expect(script).toContain("$_code = [System.Text.Encoding]::Unicode.GetString");
    expect(script).toContain("[System.Runtime.InteropServices.Marshal]::GetActiveObject");
    expect(script).toContain("Invoke-Expression $_code");
    expect(script).not.toContain("New-Object -ComObject");
  });

  test("uses PowerPoint/WPS presentation ProgIDs for presentation scripts", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      success: true,
      app: "presentation",
      progId: "PowerPoint.Application",
      engine: "powershell",
      output: "ok",
    }));

    const bridge = new OfficeScriptBridge();
    await bridge.executeScript("presentation", "$app.Name");

    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$progIds = @('PowerPoint.Application', 'Wpp.Application', 'Kwpp.Application')");
    expect(script).toContain("app = 'presentation'");
  });

  test("rejects empty scripts before calling PowerShell", async () => {
    const bridge = new OfficeScriptBridge();

    await expect(bridge.executeScript("word", "   ")).rejects.toThrow("执行 Office 脚本失败: 脚本内容不能为空");
    expect(executePowerShellMock).not.toHaveBeenCalled();
  });

  test("wraps PowerShell errors with Office script context", async () => {
    executePowerShellMock.mockRejectedValue(new Error("boom"));

    const bridge = new OfficeScriptBridge();

    await expect(bridge.executeScript("word", "$app.Name")).rejects.toThrow("执行 Office 脚本失败: boom");
  });
});
