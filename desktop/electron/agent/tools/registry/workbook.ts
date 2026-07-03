/**
 * 工作簿工具定义
 *
 * 包含工作簿检查、打开、创建、保存和切换工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 工作簿检查 */
const WORKBOOK_INSPECT_DEF: ToolDefinition = {
  name: "workbook.inspect",
  description: "获取当前打开的工作簿信息，包括工作表列表、名称、行列数。用于了解工作簿整体结构，是操作前的第一步",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

/** 打开已有工作簿 */
const WORKBOOK_OPEN_DEF: ToolDefinition = {
  name: "workbook.open",
  description: "打开指定路径的 Excel 工作簿文件（.xlsx/.xls/.csv），使其成为活动工作簿。用于操作用户指定的已有文件。打开后可用 workbook.inspect 查看内容",
  parameters: {
    type: "object",
    properties: {
      filePath: { type: "string", description: "工作簿文件的绝对路径，如 C:\\Users\\用户\\Desktop\\报表.xlsx" },
    },
    required: ["filePath"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

/** 创建新工作簿 */
const WORKBOOK_CREATE_DEF: ToolDefinition = {
  name: "workbook.create",
  description: "创建新的空白 Excel 工作簿并保存到指定路径。如果文件已存在则覆盖。创建后自动成为活动工作簿，可直接写入数据",
  parameters: {
    type: "object",
    properties: {
      filePath: { type: "string", description: "新工作簿的保存路径（绝对路径）。如果用户指定了工作文件夹，优先使用该文件夹路径。例如：C:\\Users\\用户\\工作文件夹\\新报表.xlsx" },
      sheetNames: {
        type: "array",
        description: "初始工作表名称列表（可选，默认一个 Sheet1）",
        items: { type: "string" },
      },
    },
    required: ["filePath"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

/** 保存工作簿 */
const WORKBOOK_SAVE_DEF: ToolDefinition = {
  name: "workbook.save",
  description: "保存当前活动工作簿。如果指定 saveAsPath 则另存为新文件（另存为不关闭原文件）。用于持久化修改、导出副本",
  parameters: {
    type: "object",
    properties: {
      saveAsPath: { type: "string", description: "另存为路径（可选，不填则保存到原位置）" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

/** 切换活动工作簿 */
const WORKBOOK_SWITCH_DEF: ToolDefinition = {
  name: "workbook.switch",
  description: "切换当前活动工作簿到指定名称的工作簿。用于多工作簿场景，切换后所有操作（read/write/inspect）都作用于新活动工作簿",
  parameters: {
    type: "object",
    properties: {
      workbookName: { type: "string", description: "目标工作簿名称（含扩展名），如 报表.xlsx。可用 workbook.inspect 查看所有打开的工作簿" },
    },
    required: ["workbookName"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const WORKBOOK_TOOL_DEFINITIONS: ToolDefinition[] = [
  WORKBOOK_INSPECT_DEF,
  WORKBOOK_OPEN_DEF,
  WORKBOOK_CREATE_DEF,
  WORKBOOK_SAVE_DEF,
  WORKBOOK_SWITCH_DEF,
];
