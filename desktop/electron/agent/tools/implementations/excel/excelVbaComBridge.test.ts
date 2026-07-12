import { beforeEach, describe, expect, it, vi } from "vitest";

const powershellMocks = vi.hoisted(() => ({
  executePowerShell: vi.fn(),
}));

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: powershellMocks.executePowerShell,
  };
});

import { ExcelVbaComBridge } from "./excelVbaComBridge";

function createBridge(host: "excel" | "wps" = "excel"): ExcelVbaComBridge {
  return new ExcelVbaComBridge({ host } as never);
}

function writeResult(moduleName: string, code: string) {
  return JSON.stringify({
    moduleName,
    created: true,
    lineCount: code.split(/\r?\n/).length,
    sourceBase64: Buffer.from(code, "utf16le").toString("base64"),
    saved: true,
    workbookName: "Book-macro.xlsm",
    workbookPath: "D:\\docs\\Book-macro.xlsm",
  });
}

describe("ExcelVbaComBridge", () => {
  beforeEach(() => {
    powershellMocks.executePowerShell.mockReset();
  });

  it("upserts, reads back, compiles and saves a persistent VBA module", async () => {
    const code = "Option Explicit\nPublic Sub UnitButtonClick()\nEnd Sub";
    powershellMocks.executePowerShell.mockResolvedValue(writeResult("UnitButtons", code));

    const result = await createBridge("wps").writeModule("UnitButtons", code, {
      entryPoint: "UnitButtonClick",
      save: true,
    });

    expect(result).toMatchObject({
      moduleName: "UnitButtons",
      sourceVerified: true,
      compileVerified: true,
      entryPoint: "UnitButtonClick",
      entryPointVerified: true,
      saved: true,
      host: "wps",
    });
    const script = powershellMocks.executePowerShell.mock.calls[0][0] as string;
    expect(script).toContain("VBComponents.Add(1)");
    expect(script).toContain("$excel.Run($probeMacro)");
    expect(script).toContain("VBA 模块源码回读不一致");
    expect(script).toContain("VBComponents.Remove($module)");
    expect(script).toContain("-macro.xlsm");
    expect(script).toContain("$shouldSave = $true");
  });

  it("rejects a missing entry point before changing the workbook", async () => {
    await expect(createBridge().writeModule(
      "UnitButtons",
      "Public Sub OtherMacro()\nEnd Sub",
      { entryPoint: "UnitButtonClick", save: true }
    )).rejects.toThrow("找不到入口过程: UnitButtonClick");

    expect(powershellMocks.executePowerShell).not.toHaveBeenCalled();
  });

  it("passes macro arguments through encoded PowerShell variables", async () => {
    powershellMocks.executePowerShell.mockResolvedValue(JSON.stringify({
      invoked: true,
      macroName: "UnitButtons.RenderUnit",
      returnValue: null,
    }));

    const result = await createBridge().runMacro("UnitButtons.RenderUnit", ["单位1", 2]);

    expect(result).toMatchObject({ invoked: true, macroName: "UnitButtons.RenderUnit" });
    const script = powershellMocks.executePowerShell.mock.calls[0][0] as string;
    expect(script).toContain("ConvertFrom-Json -InputObject $_argJson0");
    expect(script).toContain("$excel.Run($_macroName, $_arg0, $_arg1)");
    expect(script).not.toContain("单位1");
  });

});
