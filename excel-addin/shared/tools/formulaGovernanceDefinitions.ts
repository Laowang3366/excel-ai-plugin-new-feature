import type { ToolDefinition } from "./types";

const SCOPE = { type: "string", enum: ["workbook", "sheet", "target"] } as const;
const REPLACEMENT = {
  type: "object",
  properties: {
    find: { type: "string", minLength: 1 },
    replace: { type: "string" },
  },
  required: ["find", "replace"],
  additionalProperties: false,
} as const;

export const FORMULA_GOVERNANCE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "formula.dependencies.inspect",
    description:
      "文本解析公式依赖图（非 Excel 计算引擎）。scope=workbook|sheet|target；收集范围内真实公式单元格后构图。返回 nodes/edges/cycles/brokenReferences 与 limitations（text-parse-only 等）。不得声称引擎级循环引用。safe",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        scope: SCOPE,
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
      },
      required: ["scope"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.references.repair",
    description:
      "仅按显式 find/replace mapping 修复公式引用（无智能猜测）。先生成完整计划；仍含 #REF! 则 formula_repair_incomplete 且不写入。任何修改前写入隐藏备份表 WENGGE_FORMULA_BACKUP_V1；写后回读验证。dangerous",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      properties: {
        scope: SCOPE,
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
        replacements: {
          type: "array",
          minItems: 1,
          maxItems: 1000,
          items: REPLACEMENT,
        },
        applyAllMappings: { type: "boolean" },
      },
      required: ["scope", "replacements"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.convertToValues",
    description:
      "将目标范围内公式替换为当前计算值。必须先把公式元数据写入隐藏备份表（WENGGE_FORMULA_BACKUP_V1：backupId/createdAt/sheet/address/formula/formulaR1C1/numberFormat/locked/spillAddress/sourceRange），再写值；禁止无持久备份。写后回读验证。dangerous",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      properties: {
        scope: SCOPE,
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
        createBackup: { type: "boolean" },
        backupId: { type: "string", minLength: 1 },
      },
      required: ["scope"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.backups.inspect",
    description:
      "检查工作簿内隐藏公式备份表（magic WENGGE_FORMULA_BACKUP_V1）。校验 magic/header，容忍协议允许的坏行并报告 skippedRows，返回可恢复备份摘要。safe",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "formula.backups.restore",
    description:
      "按 backupId 从隐藏备份表恢复公式及可恢复元数据（formula/numberFormat/locked 等）；写后回读验证；默认不删除备份。backup 不存在/损坏明确失败。dangerous",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      properties: {
        backupId: { type: "string", minLength: 1 },
        removeAfterRestore: { type: "boolean" },
      },
      required: ["backupId"],
      additionalProperties: false,
    },
  },
];
