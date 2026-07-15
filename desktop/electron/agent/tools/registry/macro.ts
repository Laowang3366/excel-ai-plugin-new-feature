/**
 * Excel/WPS 内部宏工具定义
 *
 * 只包含写入办公软件内部宏工程所需的统一工具。
 */

import type { ToolDefinition } from "../../shared/types";

const MACRO_DETECT_DEF: ToolDefinition = {
  name: "macro.detect",
  description:
    "检测当前 Excel/WPS 工作簿可写入的内部宏语言。只返回 VBA 或 WPS JSA，不检测 Python、cscript 等桌面端临时执行环境",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const MACRO_RUN_DEF: ToolDefinition = {
  name: "macro.run",
  description:
    "运行当前工作簿内部已有的 VBA 宏入口并返回调用结果。WPS JSA 目前只支持可靠的写入和回读校验，不提供未经验证的远程运行入口",
  parameters: {
    type: "object",
    properties: {
      language: {
        type: "string",
        enum: ["vba"],
        description: "当前仅支持 vba",
      },
      macroName: { type: "string", description: "宏名称" },
      args: { type: "array", description: "宏参数（可选）", items: {} },
    },
    required: ["language", "macroName"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

const MACRO_WRITE_DEF: ToolDefinition = {
  name: "macro.write",
  description:
    "把代码写入当前办公软件的内部宏工程并回读校验。VBA 幂等更新指定标准模块；WPS JavaScript 更新当前 JSA 组件。此工具不在桌面助手外部执行代码，也不用于写单元格公式",
  parameters: {
    type: "object",
    properties: {
      language: {
        type: "string",
        enum: ["vba", "javascript"],
        description: "目标内部宏语言；javascript 表示 WPS JSA，不是 cscript",
      },
      moduleName: {
        type: "string",
        description: "VBA 标准模块名；JavaScript 写入当前 WPS JSA 组件时可省略",
      },
      code: { type: "string", description: "要写入内部宏工程的完整代码" },
      entryPoint: { type: "string", description: "代码中必须存在的公开入口名称" },
      saveAsPath: { type: "string", description: "仅 VBA 可用；可选的 .xlsm/.xlsb/.xls 保存路径" },
    },
    required: ["language", "code", "entryPoint"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

export const MACRO_TOOL_DEFINITIONS: ToolDefinition[] = [
  MACRO_DETECT_DEF,
  MACRO_WRITE_DEF,
  MACRO_RUN_DEF,
];
