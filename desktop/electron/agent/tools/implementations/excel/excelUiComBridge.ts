/**
 * ExcelUiComBridge — UI 控件 COM 桥接实现
 *
 * 通过 PowerShell COM 自动化操作 Excel ActiveX 控件。
 */

import type { ExcelUiBridge } from "../../contracts/excel";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type { ExcelComBridge } from "./excelComBridge";
import { ExcelVbaComBridge } from "./excelVbaComBridge";

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
    const oleClass = CONTROL_TYPE_MAP[params.controlType];
    if (!oleClass) {
      throw new Error(
        `不支持的控件类型: ${params.controlType}。支持: ${Object.keys(CONTROL_TYPE_MAP).join(", ")}`
      );
    }

    const progId = this.getProgId();
    const captionVal = params.caption || params.name;

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

    // 设置关联宏
    if (params.macroName) {
      script += `\n      ${psVar("_macroName", params.macroName)}`;
      script += `\n      $ole.OnAction = $_macroName`;
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
        $ole = $ws.OLEObjects.Item($_name)
        $ole.Delete()
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
        foreach ($ole in $ws.OLEObjects) {
          $ctrl = @{
            name = $ole.Name
            progId = $ole.OLEType
            left = $ole.Left
            top = $ole.Top
            width = $ole.Width
            height = $ole.Height
          }
          try { $ctrl.Add("caption", $ole.Object.Caption) } catch {}
          try { $ctrl.Add("linkedCell", $ole.Object.LinkedCell) } catch {}
          try { $ctrl.Add("onAction", $ole.OnAction) } catch {}
          $controls += $ctrl
        }
        $controls | ConvertTo-Json -Depth 3 -Compress
      `);
      const parsed = JSON.parse(result);
      return Array.isArray(parsed) ? parsed : [parsed];
    } catch (err: any) {
      throw new Error(`列出控件失败: ${err.message}`);
    }
  }

  async createForm(params: {
    formName: string;
    caption: string;
    width?: number;
    height?: number;
    controls?: Array<Record<string, unknown>>;
    eventCode?: string;
  }): Promise<unknown> {
    try {
      // 通过 VBA VBProject 动态创建 UserForm
      // 生成完整的 VBA 代码，包含窗体创建 + 控件添加 + 事件处理
      const vbaBridge = new ExcelVbaComBridge(this.comBridge);

      // 1. 生成创建窗体和控件的 VBA 代码
      const createFormCode = this.generateFormCreateCode(params);
      await vbaBridge.executeCode(createFormCode);

      return { success: true, formName: params.formName };
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

  /**
   * 生成创建 UserForm 的 VBA 代码
   *
   * 流程：
   * 1. 通过 VBProject.VBComponents.Add(3) 添加 MSForm 组件
   * 2. 设置窗体属性
   * 3. 在窗体上添加控件
   * 4. 添加事件处理代码
   */
  private generateFormCreateCode(params: {
    formName: string;
    caption: string;
    width?: number;
    height?: number;
    controls?: Array<Record<string, unknown>>;
    eventCode?: string;
  }): string {
    const formName = params.formName.replace(/\s+/g, "_");
    const controls = params.controls || [];
    const eventCode = params.eventCode || "";

    // 生成控件创建代码
    let controlsCode = "";
    for (const ctrl of controls) {
      const ctrlType = ctrl.type as string;
      const ctrlName = ctrl.name as string;
      const ctrlLeft = ctrl.left as number;
      const ctrlTop = ctrl.top as number;
      const ctrlWidth = ctrl.width as number;
      const ctrlHeight = ctrl.height as number;
      const ctrlCaption = (ctrl.caption as string) || ctrlName;

      // VBA MSForm 控件类型映射
      const vbaCtrlType: Record<string, string> = {
        CommandButton: "CommandButton",
        TextBox: "TextBox",
        Label: "Label",
        ComboBox: "ComboBox",
        ListBox: "ListBox",
        CheckBox: "CheckBox",
        OptionButton: "OptionButton",
        Frame: "Frame",
        SpinButton: "SpinButton",
        ScrollBar: "ScrollBar",
        Image: "Image",
        ToggleButton: "ToggleButton",
        TabStrip: "TabStrip",
        MultiPage: "MultiPage",
      };

      const vbaType = vbaCtrlType[ctrlType] || "CommandButton";
      controlsCode += `
    Set ctrl = form.Designer.Controls.Add("Forms.${vbaType}.1")
    ctrl.Name = "${ctrlName}"
    ctrl.Left = ${ctrlLeft}
    ctrl.Top = ${ctrlTop}
    ctrl.Width = ${ctrlWidth}
    ctrl.Height = ${ctrlHeight}`;
      if (ctrlCaption && ["CommandButton", "Label", "CheckBox", "OptionButton", "Frame", "ToggleButton"].includes(vbaType)) {
        controlsCode += `
    ctrl.Caption = "${ctrlCaption.replace(/"/g, '""')}"`;
      }
    }

    // 生成完整 VBA 代码
    return `Sub Main()
    Dim vbProj As Object
    Set vbProj = Application.ActiveWorkbook.VBProject

    ' 删除同名窗体（如果存在）
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = "${formName}" And comp.Type = 3 Then
            vbProj.VBComponents.Remove comp
            Exit For
        End If
    Next comp

    ' 添加 UserForm 组件 (vbext_ct_MSForm = 3)
    Dim formComp As Object
    Set formComp = vbProj.VBComponents.Add(3)
    formComp.Name = "${formName}"

    ' 设置窗体属性
    Dim form As Object
    Set form = formComp.Designer
    form.Caption = "${params.caption.replace(/"/g, '""')}"${params.width ? `\n    form.Width = ${params.width}` : ""}${params.height ? `\n    form.Height = ${params.height}` : ""}

    ' 添加控件
    Dim ctrl As Object
${controlsCode}

    ' 添加事件处理代码
    Dim codeMod As Object
    Set codeMod = formComp.CodeModule
${eventCode ? `    codeMod.AddFromString "${eventCode.replace(/"/g, '""').replace(/\n/g, '" & vbCrLf & "')}"` : ""}
End Sub`;
  }
}
