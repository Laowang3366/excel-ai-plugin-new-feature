import type { ToolDefinition } from "./types";

export const STRUCTURE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "sheet.visibility.get",
    description: "读取工作表可见性 visible|hidden|veryHidden（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.visibility.set",
    description: "设置工作表可见性 visible|hidden|veryHidden（Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        visibility: { type: "string", enum: ["visible", "hidden", "veryHidden"] },
      },
      required: ["sheetName", "visibility"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.protection.get",
    description: "读取工作表保护状态（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.protection.protect",
    description:
      "保护工作表。password 仅当前请求内存使用，不持久化/不入日志（Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        password: { type: "string" },
      },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.protection.unprotect",
    description:
      "取消工作表保护。password 仅当前请求内存使用（Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        password: { type: "string" },
      },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "namedRange.list",
    description: "列出命名区域。scope=workbook|worksheet（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        scope: { type: "string", enum: ["workbook", "worksheet"] },
        sheetName: { type: "string" },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "namedRange.create",
    description: "创建命名区域。workbook 或 worksheet scope；写后回读 name/refersTo",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", minLength: 1 },
        refersTo: { type: "string", minLength: 1 },
        scope: { type: "string", enum: ["workbook", "worksheet"] },
        sheetName: { type: "string" },
        visible: { type: "boolean" },
      },
      required: ["name", "refersTo", "scope"],
      additionalProperties: false,
    },
  },
  {
    name: "namedRange.update",
    description:
      "更新命名区域 refersTo/visible。newName 因 NamedItem.name 只读：先 add 新名（失败则旧名仍在），成功后再 delete 旧名；冲突按大小写不敏感检测",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", minLength: 1 },
        scope: { type: "string", enum: ["workbook", "worksheet"] },
        sheetName: { type: "string" },
        newName: { type: "string", minLength: 1 },
        refersTo: { type: "string", minLength: 1 },
        visible: { type: "boolean" },
      },
      required: ["name", "scope"],
      additionalProperties: false,
    },
  },
  {
    name: "namedRange.delete",
    description: "删除命名区域",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", minLength: 1 },
        scope: { type: "string", enum: ["workbook", "worksheet"] },
        sheetName: { type: "string" },
      },
      required: ["name", "scope"],
      additionalProperties: false,
    },
  },
];
