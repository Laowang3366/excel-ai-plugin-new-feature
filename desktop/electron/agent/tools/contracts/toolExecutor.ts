/** 工具风险等级 */
export type ToolRiskLevel = "safe" | "moderate" | "dangerous";

/** 工具定义 */
export interface ToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>; // JSON Schema
  riskLevel: ToolRiskLevel;
  /** 是否需要用户确认才能执行 */
  requiresApproval: boolean;
  /** 是否为文件删除类操作（range.clear / sheet.operation delete / ui.removeControl） */
  isFileDeletion?: boolean;
  /** 是否会把用户输入、查询或文件内容发送到第三方服务 */
  isDataEgress?: boolean;
  /** 即使在高权限自动模式下也必须逐次确认 */
  requiresExplicitApproval?: boolean;
  /** 工具依赖的运行时环境。未设置或 "none" 表示无特殊依赖。 */
  requiresOfficeApp?: "excel" | "word" | "presentation" | "any";
}

/** 工具执行结果 */
export interface ToolExecutionResult {
  success: boolean;
  data?: unknown;
  error?: string;
}

export interface ToolExecutionContext {
  threadId?: string;
  turnId?: string;
  userMessages: string[];
}

/** 工具执行器接口 */
export interface ToolExecutor {
  readonly name: string;
  execute(
    args: Record<string, unknown>,
    context?: ToolExecutionContext,
  ): Promise<ToolExecutionResult>;
}
