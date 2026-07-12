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

import { ExcelUiComBridge } from "./excelUiComBridge";

function createBridge(): ExcelUiComBridge {
  return new ExcelUiComBridge({ host: "wps" } as never);
}

describe("ExcelUiComBridge macro buttons", () => {
  beforeEach(() => {
    powershellMocks.executePowerShell.mockReset();
  });

  it("creates a form button and verifies its OnAction binding", async () => {
    powershellMocks.executePowerShell.mockResolvedValue(JSON.stringify({
      success: true,
      created: true,
      verified: true,
      name: "Unit1",
      controlType: "button",
      controlKind: "form",
      caption: "单位1",
      onAction: "UnitButtons.UnitButtonClick",
    }));

    const result = await createBridge().addControl({
      sheetName: "效果",
      controlType: "button",
      name: "Unit1",
      left: 10,
      top: 20,
      width: 80,
      height: 24,
      caption: "单位1",
      macroName: "UnitButtons.UnitButtonClick",
    });

    expect(result).toMatchObject({ verified: true, onAction: "UnitButtons.UnitButtonClick" });
    const script = powershellMocks.executePowerShell.mock.calls[0][0] as string;
    expect(script).toContain("$ws.Buttons().Add(10, 20, 80, 24)");
    expect(script).toContain("$button.OnAction = $_macroName");
    expect(script).toContain("按钮宏绑定回读不一致");
    expect(script).not.toContain("OLEObjects.Add");
  });

  it("does not pretend that an ActiveX control supports OnAction", async () => {
    await expect(createBridge().addControl({
      sheetName: "效果",
      controlType: "checkbox",
      name: "Check1",
      left: 10,
      top: 20,
      width: 80,
      height: 24,
      macroName: "UnitButtons.UnitButtonClick",
    })).rejects.toThrow("ActiveX 控件不能通过 OnAction 绑定宏");

    expect(powershellMocks.executePowerShell).not.toHaveBeenCalled();
  });

  it("lists form buttons with their verified macro binding", async () => {
    powershellMocks.executePowerShell.mockResolvedValue(JSON.stringify([{
      name: "Unit1",
      controlType: "button",
      controlKind: "form",
      onAction: "UnitButtons.UnitButtonClick",
    }]));

    const result = await createBridge().listControls("效果");

    expect(result).toEqual([expect.objectContaining({
      name: "Unit1",
      onAction: "UnitButtons.UnitButtonClick",
    })]);
    expect(powershellMocks.executePowerShell.mock.calls[0][0]).toContain("foreach ($button in $ws.Buttons())");
  });

  it("creates UserForms directly through VBProject without a temporary VBA module", async () => {
    powershellMocks.executePowerShell.mockResolvedValue(JSON.stringify({
      success: true,
      verified: true,
      formName: "SettingsForm",
      caption: "设置",
      controlCount: 1,
      eventCodeLines: 3,
    }));

    const result = await createBridge().createForm({
      formName: "SettingsForm",
      caption: "设置",
      controls: [{
        type: "CommandButton",
        name: "SaveButton",
        caption: "保存",
        left: 10,
        top: 10,
        width: 60,
        height: 24,
      }],
      eventCode: "Private Sub SaveButton_Click()\nEnd Sub",
    });

    expect(result).toMatchObject({ verified: true, formName: "SettingsForm" });
    const script = powershellMocks.executePowerShell.mock.calls[0][0] as string;
    expect(script).toContain("$vbProject.VBComponents.Add(3)");
    expect(script).toContain("ConvertFrom-Json -InputObject $_controlsJson");
    expect(script).toContain("$tempForm.CodeModule.AddFromString($_eventCode)");
    expect(script).toContain("WenggeOld_");
    expect(script).not.toContain("Sub Main()");
    expect(script).not.toContain("TempModule_");
  });

  it("rejects unsupported UserForm controls before changing VBProject", async () => {
    await expect(createBridge().createForm({
      formName: "SettingsForm",
      caption: "设置",
      controls: [{
        type: "UnknownControl",
        name: "BadControl",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
      }],
    })).rejects.toThrow("不支持的 UserForm 控件类型");

    expect(powershellMocks.executePowerShell).not.toHaveBeenCalled();
  });
});
