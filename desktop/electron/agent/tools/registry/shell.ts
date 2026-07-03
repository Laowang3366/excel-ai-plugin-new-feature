/**
 * Shell 工具定义
 *
 * 包含通用命令执行工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 通用 Shell 命令执行 — 参考 Codex shell_command */
const SHELL_EXECUTE_DEF: ToolDefinition = {
  name: "shell.execute",
  description:
    "在用户默认 shell 中执行命令并返回输出（Windows 用 PowerShell）。用于 Git、dir、pip、系统工具等短命令。多行 Python 或包含复杂引号的 Python 代码请用 python.execute，避免 python -c 引号失败。注意：出于安全考虑，此工具始终需要用户确认，不受权限模式影响",
  parameters: {
    type: "object",
    properties: {
      command: { type: "string", description: "要执行的命令" },
      workdir: { type: "string", description: "工作目录，默认为用户主目录" },
      timeout_ms: { type: "number", description: "超时毫秒数，默认 30000" },
    },
    required: ["command"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
};

export const SHELL_TOOL_DEFINITIONS: ToolDefinition[] = [
  SHELL_EXECUTE_DEF,
];
