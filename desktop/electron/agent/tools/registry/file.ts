/**
 * 文件上下文工具定义
 *
 * 包含常用文件系统路径查询工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 获取常用文件系统路径 */
const FILE_GET_PATHS_DEF: ToolDefinition = {
  name: "file.getPaths",
  description:
    "获取用户电脑上的常用路径，如桌面、文档、下载等目录。用于确定文件保存位置、查找文件路径。返回路径列表供后续操作使用",
  parameters: {
    type: "object",
    properties: {
      pathNames: {
        type: "array",
        description:
          "要查询的路径名称列表，如 ['desktop', 'documents', 'downloads', 'pictures', 'appData']",
        items: { type: "string" },
      },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const FILE_TOOL_DEFINITIONS: ToolDefinition[] = [FILE_GET_PATHS_DEF];
