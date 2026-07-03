/**
 * Excel UI 工具定义
 *
 * 包含工作表控件、UserForm 和菜单相关工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 添加工作表控件 */
const UI_ADD_CONTROL_DEF: ToolDefinition = {
  name: "ui.addControl",
  description: "在工作表上添加 ActiveX 控件。支持10种类型：button(按钮)、dropdown(下拉框)、checkbox(复选框)、listbox(列表框)、spinner(数值调节)、scrollbar(滚动条)、label(标签)、textbox(文本框)、optionbutton(选项按钮)、groupbox(分组框)。可通过 macroName 关联宏，linkedCell 链接单元格",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      controlType: {
        type: "string",
        enum: ["button", "dropdown", "checkbox", "listbox", "spinner", "scrollbar", "label", "textbox", "optionbutton", "groupbox"],
        description: "控件类型",
      },
      name: { type: "string", description: "控件名称" },
      left: { type: "number", description: "左边距(磅)" },
      top: { type: "number", description: "上边距(磅)" },
      width: { type: "number", description: "宽度(磅)" },
      height: { type: "number", description: "高度(磅)" },
      caption: { type: "string", description: "显示文本" },
      macroName: { type: "string", description: "关联的宏名称（按钮点击时执行）" },
      linkedCell: { type: "string", description: "链接单元格（如 A1，控件值变化时写入该单元格）" },
    },
    required: ["sheetName", "controlType", "name", "left", "top", "width", "height"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

/** 删除工作表控件 */
const UI_REMOVE_CONTROL_DEF: ToolDefinition = {
  name: "ui.removeControl",
  description: "删除工作表上的指定控件。删除前可用 ui.listControls 查看现有控件",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      name: { type: "string", description: "控件名称" },
    },
    required: ["sheetName", "name"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  isFileDeletion: true,
};

/** 列出工作表控件 */
const UI_LIST_CONTROLS_DEF: ToolDefinition = {
  name: "ui.listControls",
  description: "列出工作表上的所有控件，返回控件名称、类型、位置、大小等信息。用于了解现有控件或确认删除目标",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
    },
    required: ["sheetName"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

/** 创建 UserForm 窗体 */
const UI_CREATE_FORM_DEF: ToolDefinition = {
  name: "ui.createForm",
  description: "创建 VBA UserForm 窗体，可添加按钮、文本框、标签等控件并绑定事件处理代码。用于创建交互式对话框、数据录入窗体、设置面板等",
  parameters: {
    type: "object",
    properties: {
      formName: { type: "string", description: "窗体名称" },
      caption: { type: "string", description: "窗体标题" },
      width: { type: "number", description: "窗体宽度" },
      height: { type: "number", description: "窗体高度" },
      controls: {
        type: "array",
        description: "窗体上的控件列表",
        items: {
          type: "object",
          properties: {
            type: { type: "string", description: "控件类型: CommandButton/TextBox/Label/ComboBox/ListBox/CheckBox/OptionButton/Frame" },
            name: { type: "string", description: "控件名称" },
            caption: { type: "string", description: "显示文本" },
            left: { type: "number", description: "左边距" },
            top: { type: "number", description: "上边距" },
            width: { type: "number", description: "宽度" },
            height: { type: "number", description: "高度" },
          },
          required: ["type", "name", "left", "top", "width", "height"],
        },
      },
      eventCode: { type: "string", description: "VBA 事件处理代码（如按钮点击事件）" },
    },
    required: ["formName", "caption"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

/** 添加自定义菜单项 */
const UI_ADD_MENU_DEF: ToolDefinition = {
  name: "ui.addMenu",
  description: "在 Excel 菜单栏、右键菜单或工具栏中添加自定义菜单项，点击时执行指定宏。用于创建快捷操作入口、自定义工具菜单",
  parameters: {
    type: "object",
    properties: {
      menuBar: { type: "string", enum: ["worksheet", "cell", "toolbar"], description: "菜单位置：worksheet(工作表菜单栏)、cell(单元格右键菜单)、toolbar(工具栏)" },
      caption: { type: "string", description: "菜单项显示文本" },
      macroName: { type: "string", description: "点击时执行的宏名称" },
      beforeId: { type: "number", description: "插入位置（可选）" },
      faceId: { type: "number", description: "图标ID（可选）" },
    },
    required: ["menuBar", "caption", "macroName"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

export const UI_TOOL_DEFINITIONS: ToolDefinition[] = [
  UI_ADD_CONTROL_DEF,
  UI_REMOVE_CONTROL_DEF,
  UI_LIST_CONTROLS_DEF,
  UI_CREATE_FORM_DEF,
  UI_ADD_MENU_DEF,
];
