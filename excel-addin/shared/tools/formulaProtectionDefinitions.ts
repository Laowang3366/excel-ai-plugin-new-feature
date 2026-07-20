import type { ToolDefinition } from "./types";

export const FORMULA_PROTECTION_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "formula.protection.inspect",
    description:
      "扫描范围内公式单元格的锁定状态与工作表保护（ExcelApi 1.2 Range.format.protection.locked）。scope=workbook|sheet|target；sheet/target 需 sheetName；target 需 range。返回 formulaCount/lockedFormulaCount/sheetProtected/limitations。不返回密码。WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        scope: { type: "string", enum: ["workbook", "sheet", "target"] },
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
      },
      required: ["scope"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.protection.manage",
    description:
      "对范围内**公式单元格** lock/unlock（非整表冒充）。command=lock|unlock；scope 同 inspect。lock 默认 unlockInputs=true（仅目标范围内先解锁再锁公式，使输入格可编辑）且 protectSheet=true。password 仅当前请求内存，禁止写入结果/日志/持久化。写后回读校验。dangerous。WPS unsupported",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      properties: {
        command: { type: "string", enum: ["lock", "unlock"] },
        scope: { type: "string", enum: ["workbook", "sheet", "target"] },
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
        password: { type: "string", maxLength: 255 },
        unlockInputs: { type: "boolean" },
        protectSheet: { type: "boolean" },
      },
      required: ["command", "scope"],
      additionalProperties: false,
    },
  },
];
