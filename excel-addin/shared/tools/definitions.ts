import type { ToolDefinition } from "./types";
import { CHART_TOOL_DEFINITIONS } from "./chartDefinitions";
import { CHART_SERIES_TOOL_DEFINITIONS } from "./chartSeriesDefinitions";
import { CHART_SOURCE_TOOL_DEFINITIONS } from "./chartSourceDefinitions";
import { CHART_AXES_TOOL_DEFINITIONS } from "./chartAxesDefinitions";
import { CHART_DATA_LABELS_TOOL_DEFINITIONS } from "./chartDataLabelsDefinitions";
import { CHART_SERIES_AXIS_GROUP_TOOL_DEFINITIONS } from "./chartSeriesAxisGroupDefinitions";
import { CHART_SERIES_ADD_TOOL_DEFINITIONS } from "./chartSeriesAddDefinitions";
import { CHART_SERIES_DELETE_TOOL_DEFINITIONS } from "./chartSeriesDeleteDefinitions";
import { CHART_SERIES_VALUES_TOOL_DEFINITIONS } from "./chartSeriesValuesDefinitions";
import { CHART_SERIES_BUBBLE_SIZES_TOOL_DEFINITIONS } from "./chartSeriesBubbleSizesDefinitions";
import { CHART_SERIES_TRENDLINE_TOOL_DEFINITIONS } from "./chartSeriesTrendlineDefinitions";
import { CHART_IMAGE_TOOL_DEFINITIONS } from "./chartImageDefinitions";
import { RANGE_IMAGE_TOOL_DEFINITIONS } from "./rangeImageDefinitions";
import { RANGE_STRUCTURE_TOOL_DEFINITIONS } from "./rangeStructureDefinitions";
import { DISPLAY_TOOL_DEFINITIONS } from "./displayDefinitions";
import { FREEZE_TOOL_DEFINITIONS } from "./freezeDefinitions";
import { OBJECT_UPDATE_TOOL_DEFINITIONS } from "./objectUpdateDefinitions";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "./pageLayoutDefinitions";
import { SHAPE_TOOL_DEFINITIONS } from "./shapeDefinitions";
import { STRUCTURE_TOOL_DEFINITIONS } from "./structureDefinitions";
import { TABLE_UNLIST_TOOL_DEFINITIONS } from "./tableUnlistDefinitions";
import { TABLE_FILTER_TOOL_DEFINITIONS } from "./tableFilterDefinitions";
import { TABLE_SORT_TOOL_DEFINITIONS } from "./tableSortDefinitions";
import { FORMULA_PROTECTION_TOOL_DEFINITIONS } from "./formulaProtectionDefinitions";
import { FORMULA_GOVERNANCE_TOOL_DEFINITIONS } from "./formulaGovernanceDefinitions";
import {
  CONDITIONAL_FORMAT_TOOL_DEFINITIONS,
  DATA_VALIDATION_TOOL_DEFINITIONS,
} from "./validationDefinitions";
import { PIVOT_TOOL_DEFINITIONS } from "./pivotDefinitions";

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "host.status",
    description: "读取当前宿主连接状态与工作簿名称",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    name: "selection.get",
    description: "读取当前选区地址、值与公式",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    name: "range.read",
    description:
      '读取指定工作表区域的值与公式。可选 expand: none|spill|currentArray|currentRegion。省略 expand 且为单单元格时与桌面一致自动探测 spill；显式 expand:"none" 强制不扩展。WPS 上非 none expand 为 unsupported。',
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        expand: {
          type: "string",
          enum: ["none", "spill", "currentArray", "currentRegion"],
        },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "range.write",
    description: "写入区域值；可选 verify 触发写后回读",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        values: { type: "array" },
        verify: { type: "boolean" },
      },
      required: ["sheetName", "range", "values"],
      additionalProperties: false,
    },
  },
  {
    name: "range.clear",
    description: "清除指定区域内容",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "range.format.read",
    description: "读取区域格式（字体、填充、数字格式、对齐、换行）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "range.format.write",
    description: "写入区域格式字段（按宿主能力）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        format: {
          type: "object",
          properties: {
            fontName: { type: "string" },
            fontSize: { type: "number" },
            fontBold: { type: "boolean" },
            fontColor: { type: "string" },
            fillColor: { type: "string" },
            numberFormat: { type: "string" },
            horizontalAlignment: { type: "string" },
            verticalAlignment: { type: "string" },
            wrapText: { type: "boolean" },
          },
          additionalProperties: false,
        },
      },
      required: ["sheetName", "range", "format"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.read",
    description: "读取区域公式（基于 range.read）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.write",
    description: "写入单个锚点公式；可选 verify 写后回读",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        formula: { type: "string" },
        verify: { type: "boolean" },
      },
      required: ["sheetName", "range", "formula"],
      additionalProperties: false,
    },
  },
  {
    name: "formula.context",
    description:
      "读取区域内含公式单元格。返回 { sheetName, address, formulas:[{address,formula,value}] }。range 可省略，省略时使用 UsedRange（与桌面一致）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.list",
    description: "列出工作表",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    name: "sheet.operation",
    description:
      "统一工作表操作：add|rename|delete|copy|move。position 为 1-based（与桌面 COM 合同一致）。WPS 上 copy/move 为 unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        operation: {
          type: "string",
          enum: ["add", "rename", "delete", "copy", "move"],
        },
        sheetName: { type: "string" },
        newName: { type: "string" },
        position: { type: "integer", minimum: 1 },
      },
      required: ["operation", "sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.add",
    description: "新增工作表",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.rename",
    description: "重命名工作表",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        newName: { type: "string" },
      },
      required: ["sheetName", "newName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.delete",
    description: "删除工作表",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "table.list",
    description: "列出表格（可选按工作表过滤）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "table.create",
    description: "从区域创建基础表格",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        name: { type: "string" },
        hasHeaders: { type: "boolean" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "table.delete",
    description: "删除指定表格",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        tableName: { type: "string" },
      },
      required: ["sheetName", "tableName"],
      additionalProperties: false,
    },
  },
  {
    name: "workbook.inspect",
    description: "检查工作簿/活动表 used range；Office.js 每表含 usedRangeAddress/rowCount/columnCount",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    name: "workbook.objects.inspect",
    description:
      "只读聚合当前工作簿对象清单：sheets + tables/charts/namedRanges/shapes 分类。maxItemsPerCategory 默认 100（1..500），超限 truncated=true 且保留真实 totalCount；可选 sheetName 过滤表级对象（workbook 命名区域仍包含）。单分类 unsupported/failed 不拖垮整项；WPS 上 table/chart/shape 为 typed unsupported。",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        maxItemsPerCategory: { type: "integer", minimum: 1, maximum: 500 },
        sheetName: { type: "string", minLength: 1 },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "workbook.save",
    description:
      "保存当前宿主已打开的工作簿（原地 Save）。无路径参数；不支持 saveAs/打开/创建/切换工作簿。未命名新簿可能弹出宿主另存对话框或失败。Office.js ExcelApi 1.1；WPS 需 ActiveWorkbook.Save 成员。",
    riskLevel: "moderate",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  ...CONDITIONAL_FORMAT_TOOL_DEFINITIONS,
  ...DATA_VALIDATION_TOOL_DEFINITIONS,
  ...CHART_TOOL_DEFINITIONS,
  ...CHART_SERIES_TOOL_DEFINITIONS,
  ...CHART_SOURCE_TOOL_DEFINITIONS,
  ...CHART_AXES_TOOL_DEFINITIONS,
  ...CHART_DATA_LABELS_TOOL_DEFINITIONS,
  ...CHART_SERIES_AXIS_GROUP_TOOL_DEFINITIONS,
  ...CHART_SERIES_DELETE_TOOL_DEFINITIONS,
  ...CHART_SERIES_ADD_TOOL_DEFINITIONS,
  ...CHART_SERIES_VALUES_TOOL_DEFINITIONS,
  ...CHART_SERIES_BUBBLE_SIZES_TOOL_DEFINITIONS,
  ...CHART_SERIES_TRENDLINE_TOOL_DEFINITIONS,
  ...CHART_IMAGE_TOOL_DEFINITIONS,
  ...RANGE_IMAGE_TOOL_DEFINITIONS,
  ...RANGE_STRUCTURE_TOOL_DEFINITIONS,
  ...STRUCTURE_TOOL_DEFINITIONS,
  ...OBJECT_UPDATE_TOOL_DEFINITIONS,
  ...DISPLAY_TOOL_DEFINITIONS,
  ...FREEZE_TOOL_DEFINITIONS,
  ...PAGE_LAYOUT_TOOL_DEFINITIONS,
  ...SHAPE_TOOL_DEFINITIONS,
  ...TABLE_UNLIST_TOOL_DEFINITIONS,
  ...TABLE_FILTER_TOOL_DEFINITIONS,
  ...TABLE_SORT_TOOL_DEFINITIONS,
  ...FORMULA_PROTECTION_TOOL_DEFINITIONS,
  ...FORMULA_GOVERNANCE_TOOL_DEFINITIONS,
  ...PIVOT_TOOL_DEFINITIONS,
];
