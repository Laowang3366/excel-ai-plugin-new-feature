/**
 * Python 工具定义
 *
 * 面向模型暴露通用 Python 脚本执行能力，避免 shell.execute 中的多层引号转义问题。
 */

import type { ToolDefinition } from "../../shared/types";

const PYTHON_EXECUTE_DEF: ToolDefinition = {
  name: "python.execute",
  description:
    "执行通用 Python 脚本并返回 stdout/stderr/exitCode。适合文件级处理、Word/Excel/PPT 文件脚本、数据转换等；不要把长 Python 代码塞进 shell.execute 的 python -c",
  parameters: {
    type: "object",
    properties: {
      code: { type: "string", description: "完整 Python 脚本代码" },
      workdir: { type: "string", description: "工作目录，默认用户主目录；非白名单路径会重定向到临时目录" },
      timeout_ms: { type: "number", description: "超时毫秒数，默认 90000" },
    },
    required: ["code"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

export const PYTHON_TOOL_DEFINITIONS: ToolDefinition[] = [
  PYTHON_EXECUTE_DEF,
];
