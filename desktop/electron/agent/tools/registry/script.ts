/**
 * Excel 脚本工具定义
 *
 * 包含 VBA 和统一脚本执行相关工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 执行 VBA 宏 */
const VBA_RUN_MACRO_DEF: ToolDefinition = {
  name: "vba.runMacro",
  description: "运行工作簿中已有的 VBA 宏。仅用于执行已存在的宏，一次性脚本请用 script.execute",
  parameters: {
    type: "object",
    properties: {
      macroName: { type: "string", description: "宏名称" },
      args: { type: "array", description: "宏参数（可选）", items: {} },
    },
    required: ["macroName"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

/** 写入 VBA 模块 */
const VBA_WRITE_MODULE_DEF: ToolDefinition = {
  name: "vba.writeModule",
  description: "向工作簿持久化写入 VBA 代码模块，代码会保留在工作簿中。用于创建可重复使用的宏、事件处理模块。一次性执行请用 script.execute",
  parameters: {
    type: "object",
    properties: {
      moduleName: { type: "string", description: "模块名称" },
      code: { type: "string", description: "VBA 代码" },
    },
    required: ["moduleName", "code"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

/** 脚本环境检测 */
const SCRIPT_DETECT_DEF: ToolDefinition = {
  name: "script.detect",
  description: "检测当前环境支持的脚本语言及推荐优先级。WPS 优先 JS(cscript)，Office Excel 优先 VBA。返回可用语言列表、引擎类型和推荐语言",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

/** 统一脚本执行 */
const SCRIPT_EXECUTE_DEF: ToolDefinition = {
  name: "script.execute",
  description: "执行脚本代码，自动根据环境选择最优语言（WPS优先JS，Excel优先VBA），失败自动切换备选语言。用于自动化、批量操作、复杂逻辑等场景。指定 language 可强制使用某语言",
  parameters: {
    type: "object",
    properties: {
      code: { type: "string", description: "代码内容" },
      language: {
        type: "string",
        enum: ["vba", "javascript", "python"],
        description: "指定语言（不填则自动选择推荐语言）",
      },
    },
    required: ["code"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

export const SCRIPT_TOOL_DEFINITIONS: ToolDefinition[] = [
  VBA_RUN_MACRO_DEF,
  VBA_WRITE_MODULE_DEF,
  SCRIPT_DETECT_DEF,
  SCRIPT_EXECUTE_DEF,
];
