import { withOfficeOperationDiscriminator, WORKFLOW_OPERATIONS } from "./officeActionSchemas";

export const OFFICE_WORKFLOW_VARIABLES_SCHEMA: Record<string, unknown> = {
  type: "object",
  maxProperties: 128,
  propertyNames: {
    type: "string",
    pattern: "^[A-Za-z_][A-Za-z0-9_-]{0,63}$",
  },
  additionalProperties: true,
  description:
    "模板变量，最多 128 个顶层键；键名仅允许字母或下划线开头，后续使用字母、数字、下划线或连字符。嵌套对象仍可通过 {{vars.customer.name}} 引用。",
};

export const OFFICE_WORKFLOW_STEP_SCHEMA: Record<string, unknown> = withOfficeOperationDiscriminator({
  type: "object",
  properties: {
    app: { type: "string", enum: ["excel", "word", "presentation"] },
    action: {
      type: "string",
      enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"],
    },
    operation: { type: "string" },
    filePath: { type: "string" },
    outputPath: { type: "string" },
    target: { type: "string" },
    preferEngine: { type: "string", enum: ["openxml", "com"] },
    params: {
      type: "object",
      description:
        "高级 Excel 步骤必须声明语义边界：Power Query 用 advancedIntent:'refreshable-etl'，创建/更新另需 sourceKind:'external'|'multi-source'；透视表/切片器用 advancedIntent:'interactive-pivot'",
    },
    id: { type: "string", description: "可选稳定步骤 ID，供占位符和条件引用" },
    parallelGroup: { type: "string", description: "连续且同名的步骤并行执行；写入目标不得重复" },
    timeoutMs: {
      type: "integer",
      minimum: 5_000,
      maximum: 600_000,
      description: "COM 步骤超时，5000-600000 毫秒",
    },
    retry: {
      type: "object",
      properties: {
        maxAttempts: { type: "integer", minimum: 1, maximum: 5, description: "最大尝试次数，1-5" },
        delayMs: {
          type: "integer",
          minimum: 0,
          maximum: 10_000,
          description: "重试基础等待，0-10000 毫秒",
        },
      },
    },
    when: {
      type: "object",
      properties: {
        step: { description: "依赖步骤序号或 ID" },
        status: { type: "string", enum: ["done", "failed", "skipped"] },
        dataPath: { type: "string", description: "可选结果 data 路径" },
        equals: { description: "期望值" },
        exists: { type: "boolean" },
      },
      required: ["step"],
    },
  },
  required: ["app", "action", "operation", "filePath"],
}, WORKFLOW_OPERATIONS);

export const OFFICE_WORKFLOW_STEPS_SCHEMA: Record<string, unknown> = {
  type: "array",
  minItems: 1,
  maxItems: 20,
  items: OFFICE_WORKFLOW_STEP_SCHEMA,
};
