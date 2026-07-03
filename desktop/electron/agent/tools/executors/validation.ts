/**
 * 工具参数校验
 *
 * 被各领域执行器复用，只负责校验必填参数和基础类型，不承载业务兜底逻辑。
 */

export type RequiredArgType = "string" | "number" | "array" | "object";

/**
 * 校验工具参数，返回缺失或类型错误的字段描述。
 */
export function validateArgs(
  args: Record<string, unknown>,
  required: Record<string, RequiredArgType>
): string | null {
  for (const [key, expectedType] of Object.entries(required)) {
    const val = args[key];
    if (val === undefined || val === null) {
      return `缺少必填参数: ${key}`;
    }
    switch (expectedType) {
      case "string":
        if (typeof val !== "string") return `参数 ${key} 应为字符串，实际为 ${typeof val}`;
        break;
      case "number":
        if (typeof val !== "number") return `参数 ${key} 应为数字，实际为 ${typeof val}`;
        break;
      case "array":
        if (!Array.isArray(val)) return `参数 ${key} 应为数组，实际为 ${typeof val}`;
        break;
      case "object":
        if (typeof val !== "object" || Array.isArray(val)) return `参数 ${key} 应为对象，实际为 ${Array.isArray(val) ? "数组" : typeof val}`;
        break;
    }
  }
  return null;
}
