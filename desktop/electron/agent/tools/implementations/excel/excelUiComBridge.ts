/**
 * ExcelUiComBridge — UI 控件 COM 桥接实现
 *
 * 通过 PowerShell COM 自动化操作 Excel ActiveX 控件。
 */

import type { ExcelUiBridge } from "../../contracts/excel";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type { ExcelComBridge } from "./excelComBridge";

/** ActiveX 控件类型映射 */
const CONTROL_TYPE_MAP: Record<string, string> = {
  button: "Forms.CommandButton.1",
  dropdown: "Forms.ComboBox.1",
  checkbox: "Forms.CheckBox.1",
  textbox: "Forms.TextBox.1",
  label: "Forms.Label.1",
  listbox: "Forms.ListBox.1",
  spinner: "Forms.SpinButton.1",
  scrollbar: "Forms.ScrollBar.1",
  optionbutton: "Forms.OptionButton.1",
  groupbox: "Forms.Frame.1",
};

export class ExcelUiComBridge implements ExcelUiBridge {
  private comBridge: ExcelComBridge;

  constructor(comBridge: ExcelComBridge) {
    this.comBridge = comBridge;
  }

  /** 获取 COM ProgID */
  private getProgId(): string {
    return this.comBridge.host === "wps" ? "Ket.Application" : "Excel.Application";
  }

  async addControl(params: {
    sheetName: string;
    controlType: string;
    name: string;
    left: number;
    top: number;
    width: number;
    height: number;
    caption?: string;
    macroName?: string;
    linkedCell?: string;
  }): Promise<unknown> {
    if (params.controlType === "button") {
      return this.addFormButton(params);
    }

    const oleClass = CONTROL_TYPE_MAP[params.controlType];
    if (!oleClass) {
      throw new Error(
        `不支持的控件类型: ${params.controlType}。支持: ${Object.keys(CONTROL_TYPE_MAP).join(", ")}`
      );
    }

    const progId = this.getProgId();
    const captionVal = params.caption || params.name;

    if (params.macroName) {
      throw new Error("ActiveX 控件不能通过 OnAction 绑定宏；需要宏回调时请使用 button 窗体按钮");
    }

    // 构建 PowerShell 脚本
    let script = `
      ${psVar("_sheetName", params.sheetName)}
      ${psVar("_name", params.name)}
      ${psVar("_caption", captionVal)}
      $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
      $ws = $excel.ActiveWorkbook.Sheets.Item($_sheetName)
      $ole = $ws.OLEObjects.Add(ClassType="${oleClass}")
      $ole.Name = $_name
      $ole.Left = ${params.left}
      $ole.Top = ${params.top}
      $ole.Width = ${params.width}
      $ole.Height = ${params.height}
    `;

    // 设置控件标题（部分控件支持）
    if (captionVal) {
      // CommandButton, Label, CheckBox, OptionButton, Frame 支持 Caption
      if (["button", "label", "checkbox", "optionbutton", "groupbox"].includes(params.controlType)) {
        script += `\n      $ole.Object.Caption = $_caption`;
      }
      // ComboBox, TextBox 不设 Caption
    }

    // 设置链接单元格
    if (params.linkedCell) {
      script += `\n      ${psVar("_linkedCell", params.linkedCell)}`;
      script += `\n      $ole.Object.LinkedCell = $_linkedCell`;
    }

    try {
      await executePowerShell(script);
      return { success: true, name: params.name, controlType: params.controlType };
    } catch (err: any) {
      throw new Error(`添加控件失败: ${err.message}`);
    }
  }

  async removeControl(sheetName: string, name: string): Promise<void> {
    const progId = this.getProgId();
    try {
      await executePowerShell(`
        ${psVar("_sheetName", sheetName)}
        ${psVar("_name", name)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $ws = $excel.ActiveWorkbook.Sheets.Item($_sheetName)
        $removed = $false
        try {
          $button = $ws.Buttons().Item($_name)
          $button.Delete()
          $removed = $true
        } catch {}
        if (-not $removed) {
          try {
            $ole = $ws.OLEObjects.Item($_name)
            $ole.Delete()
            $removed = $true
          } catch {}
        }
        if (-not $removed) { throw "找不到控件: $_name" }
      `);
    } catch (err: any) {
      throw new Error(`删除控件失败: ${err.message}`);
    }
  }

  async listControls(sheetName: string): Promise<unknown[]> {
    const progId = this.getProgId();
    try {
      const result = await executePowerShell(`
        ${psVar("_sheetName", sheetName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $ws = $excel.ActiveWorkbook.Sheets.Item($_sheetName)
        $controls = @()
        foreach ($button in $ws.Buttons()) {
          $controls += [PSCustomObject]@{
            name = $button.Name
            controlType = "button"
            controlKind = "form"
            left = $button.Left
            top = $button.Top
            width = $button.Width
            height = $button.Height
            caption = $button.Caption
            onAction = $button.OnAction
          }
        }
        foreach ($ole in $ws.OLEObjects) {
          $ctrl = @{
            name = $ole.Name
            controlType = "activex"
            controlKind = "activex"
            progId = $ole.OLEType
            left = $ole.Left
            top = $ole.Top
            width = $ole.Width
            height = $ole.Height
          }
          try { $ctrl.Add("caption", $ole.Object.Caption) } catch {}
          try { $ctrl.Add("linkedCell", $ole.Object.LinkedCell) } catch {}
          $controls += $ctrl
        }
        ConvertTo-Json -InputObject @($controls) -Depth 3 -Compress
      `);
      if (!result) return [];
      const parsed = JSON.parse(result);
      return Array.isArray(parsed) ? parsed : [parsed];
    } catch (err: any) {
      throw new Error(`列出控件失败: ${err.message}`);
    }
  }

  private async addFormButton(params: {
    sheetName: string;
    name: string;
    left: number;
    top: number;
    width: number;
    height: number;
    caption?: string;
    macroName?: string;
  }): Promise<unknown> {
    const progId = this.getProgId();
    const caption = params.caption || params.name;
    const result = await executePowerShell(`
      ${psVar("_sheetName", params.sheetName)}
      ${psVar("_name", params.name)}
      ${psVar("_caption", caption)}
      ${psVar("_macroName", params.macroName ?? "")}
      $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
      $ws = $excel.ActiveWorkbook.Sheets.Item($_sheetName)
      $button = $null
      $created = $false
      try { $button = $ws.Buttons().Item($_name) } catch {}
      if ($null -eq $button) {
        $button = $ws.Buttons().Add(${params.left}, ${params.top}, ${params.width}, ${params.height})
        $button.Name = $_name
        $created = $true
      }
      $button.Left = ${params.left}
      $button.Top = ${params.top}
      $button.Width = ${params.width}
      $button.Height = ${params.height}
      $button.Caption = $_caption
      if ($_macroName) { $button.OnAction = $_macroName }

      $verified = $ws.Buttons().Item($_name)
      if ($verified.Caption -ne $_caption) { throw "按钮标题回读不一致" }
      if ($_macroName -and $verified.OnAction -ne $_macroName) { throw "按钮宏绑定回读不一致" }
      [PSCustomObject]@{
        success = $true
        created = $created
        verified = $true
        name = $verified.Name
        controlType = "button"
        controlKind = "form"
        caption = $verified.Caption
        onAction = $verified.OnAction
      } | ConvertTo-Json -Compress
    `);
    return JSON.parse(result);
  }

  async createForm(params: {
    formName: string;
    caption: string;
    width?: number;
    height?: number;
    controls?: Array<Record<string, unknown>>;
    eventCode?: string;
  }): Promise<unknown> {
    const formName = params.formName.trim().replace(/\s+/g, "_");
    if (!/^[A-Za-z_][A-Za-z0-9_]{0,30}$/.test(formName)) {
      throw new Error("窗体名称必须是 1 到 31 位的 VBA 标识符");
    }
    const captionTypes = new Set([
      "CommandButton", "Label", "CheckBox", "OptionButton", "Frame", "ToggleButton",
    ]);
    const controlTypeMap: Record<string, string> = {
      CommandButton: "Forms.CommandButton.1",
      TextBox: "Forms.TextBox.1",
      Label: "Forms.Label.1",
      ComboBox: "Forms.ComboBox.1",
      ListBox: "Forms.ListBox.1",
      CheckBox: "Forms.CheckBox.1",
      OptionButton: "Forms.OptionButton.1",
      Frame: "Forms.Frame.1",
      SpinButton: "Forms.SpinButton.1",
      ScrollBar: "Forms.ScrollBar.1",
      Image: "Forms.Image.1",
      ToggleButton: "Forms.ToggleButton.1",
      TabStrip: "Forms.TabStrip.1",
      MultiPage: "Forms.MultiPage.1",
    };
    const names = new Set<string>();
    const controls = (params.controls ?? []).map((control) => {
      const type = typeof control.type === "string" ? control.type : "";
      const name = typeof control.name === "string" ? control.name.trim() : "";
      if (!controlTypeMap[type]) throw new Error(`不支持的 UserForm 控件类型: ${type || "空"}`);
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) throw new Error(`无效的控件名称: ${name || "空"}`);
      if (names.has(name)) throw new Error(`控件名称重复: ${name}`);
      names.add(name);
      const dimensions = ["left", "top", "width", "height"] as const;
      for (const key of dimensions) {
        if (typeof control[key] !== "number" || !Number.isFinite(control[key])) {
          throw new Error(`控件 ${name} 的 ${key} 必须是有限数字`);
        }
      }
      return {
        progId: controlTypeMap[type],
        name,
        left: control.left,
        top: control.top,
        width: control.width,
        height: control.height,
        hasCaption: captionTypes.has(type),
        caption: typeof control.caption === "string" ? control.caption : name,
      };
    });

    try {
      const progId = this.getProgId();
      const result = await executePowerShell(`
        ${psVar("_formName", formName)}
        ${psVar("_caption", params.caption)}
        ${psVar("_controlsJson", JSON.stringify(controls))}
        ${psVar("_eventCode", params.eventCode ?? "")}
        $requestedWidth = ${params.width === undefined ? "$null" : params.width}
        $requestedHeight = ${params.height === undefined ? "$null" : params.height}

        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { throw "当前没有活动工作簿" }
        try {
          $vbProject = $wb.VBProject
          $null = $vbProject.VBComponents.Count
        } catch {
          throw "无法访问 VBA 工程，请开启‘信任对 VBA 工程对象模型的访问’"
        }

        $oldForm = $null
        foreach ($component in $vbProject.VBComponents) {
          if ($component.Name -eq $_formName) {
            if ($component.Type -ne 3) { throw "同名 VBA 组件不是 UserForm: $_formName" }
            $oldForm = $component
            break
          }
        }

        $tempForm = $null
        $backupName = ""
        try {
          $tempForm = $vbProject.VBComponents.Add(3)
          $tempForm.Name = "WenggeForm_" + [Guid]::NewGuid().ToString("N").Substring(0, 12)
          $designer = $tempForm.Designer
          $designer.Caption = $_caption
          if ($null -ne $requestedWidth) { $designer.Width = $requestedWidth }
          if ($null -ne $requestedHeight) { $designer.Height = $requestedHeight }

          $controls = ConvertFrom-Json -InputObject $_controlsJson
          if ($null -eq $controls) { $controls = @() }
          foreach ($definition in @($controls)) {
            $control = $designer.Controls.Add([string]$definition.progId)
            $control.Name = [string]$definition.name
            $control.Left = [double]$definition.left
            $control.Top = [double]$definition.top
            $control.Width = [double]$definition.width
            $control.Height = [double]$definition.height
            if ([bool]$definition.hasCaption) { $control.Caption = [string]$definition.caption }
          }
          if ($_eventCode) { $tempForm.CodeModule.AddFromString($_eventCode) }

          if ($designer.Caption -ne $_caption) { throw "UserForm 标题回读不一致" }
          if ($designer.Controls.Count -ne @($controls).Count) { throw "UserForm 控件数量回读不一致" }
          foreach ($definition in @($controls)) {
            $verified = $designer.Controls.Item([string]$definition.name)
            if ($null -eq $verified) { throw "UserForm 控件回读失败: $($definition.name)" }
          }

          if ($null -ne $oldForm) {
            $backupName = "WenggeOld_" + [Guid]::NewGuid().ToString("N").Substring(0, 12)
            $oldForm.Name = $backupName
          }
          $tempForm.Name = $_formName
          if ($null -ne $oldForm) { $vbProject.VBComponents.Remove($oldForm) }

          [PSCustomObject]@{
            success = $true
            verified = $true
            formName = $tempForm.Name
            caption = $tempForm.Designer.Caption
            controlCount = $tempForm.Designer.Controls.Count
            eventCodeLines = $tempForm.CodeModule.CountOfLines
          } | ConvertTo-Json -Compress
        } catch {
          if ($null -ne $tempForm) {
            try { $vbProject.VBComponents.Remove($tempForm) } catch {}
          }
          if ($null -ne $oldForm -and $backupName) {
            try { $oldForm.Name = $_formName } catch {}
          }
          throw
        }
      `);
      return JSON.parse(result);
    } catch (err: any) {
      throw new Error(`创建窗体失败: ${err.message}`);
    }
  }

  async addMenu(params: {
    menuBar: string;
    caption: string;
    macroName: string;
    beforeId?: number;
    faceId?: number;
  }): Promise<unknown> {
    const progId = this.getProgId();
    // 将菜单栏名称映射为 CommandBars 索引
    const commandBarName = params.menuBar === "cell"
      ? "Cell"
      : params.menuBar === "toolbar"
      ? "Standard"
      : "Worksheet Menu Bar";

    try {
      let script = `
        ${psVar("_caption", params.caption)}
        ${psVar("_macro", params.macroName)}
        ${psVar("_barName", commandBarName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $cb = $excel.CommandBars.Item($_barName)
        $btn = $cb.Controls.Add(Type:=1)
        $btn.Caption = $_caption
        $btn.OnAction = $_macro
        $btn.BeginGroup = $true
      `;

      if (params.faceId) {
        script += `\n        $btn.FaceId = ${params.faceId}`;
      }

      await executePowerShell(script);
      return { success: true, caption: params.caption, menuBar: params.menuBar };
    } catch (err: any) {
      throw new Error(`添加菜单失败: ${err.message}`);
    }
  }

}
