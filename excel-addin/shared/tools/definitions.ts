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
import { CHART_IMAGE_TOOL_DEFINITIONS } from "./chartImageDefinitions";
import { DISPLAY_TOOL_DEFINITIONS } from "./displayDefinitions";
import { FREEZE_TOOL_DEFINITIONS } from "./freezeDefinitions";
import { OBJECT_UPDATE_TOOL_DEFINITIONS } from "./objectUpdateDefinitions";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "./pageLayoutDefinitions";
import { SHAPE_TOOL_DEFINITIONS } from "./shapeDefinitions";
import { STRUCTURE_TOOL_DEFINITIONS } from "./structureDefinitions";
import { TABLE_UNLIST_TOOL_DEFINITIONS } from "./tableUnlistDefinitions";

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "host.status",
    description: "读取当前宿主连接状态与工作簿名称",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [] },
  },
  {
    name: "selection.get",
    description: "读取当前选区地址、值与公式",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [] },
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
        format: { type: "object" },
      },
      required: ["sheetName", "range", "format"],
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
    },
  },
  {
    name: "sheet.list",
    description: "列出工作表",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [] },
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
    },
  },
  {
    name: "workbook.inspect",
    description: "检查工作簿/活动表 used range；Office.js 每表含 usedRangeAddress/rowCount/columnCount",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, required: [] },
  },
  {
    name: "conditionalFormat.list",
    description: "列出区域条件格式（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
    },
  },
  {
    name: "conditionalFormat.add",
    description:
      "添加条件格式。rule.kind=cellValue|custom；cellValue 需 operator/formula1；custom 需 formula",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        rule: {
          type: "object",
          properties: {
            kind: { type: "string", enum: ["cellValue", "custom"] },
            operator: {
              type: "string",
              enum: ["greaterThan", "lessThan", "equalTo", "between", "notBetween"],
            },
            formula1: { type: "string" },
            formula2: { type: "string" },
            formula: { type: "string" },
            fillColor: { type: "string" },
            fontColor: { type: "string" },
          },
          required: ["kind"],
          additionalProperties: false,
        },
      },
      required: ["sheetName", "range", "rule"],
    },
  },
  {
    name: "conditionalFormat.delete",
    description: "按 id 删除区域条件格式",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        id: { type: "string" },
      },
      required: ["sheetName", "range", "id"],
    },
  },
  {
    name: "dataValidation.read",
    description: "读取区域数据验证规则（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
    },
  },
  {
    name: "dataValidation.write",
    description:
      "写入数据验证。rule.type=list|wholeNumber；list 用 listValues；wholeNumber 用 operator/formula1/formula2",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        rule: {
          type: "object",
          properties: {
            type: { type: "string", enum: ["list", "wholeNumber"] },
            operator: {
              type: "string",
              enum: ["between", "notBetween", "equalTo", "greaterThan", "lessThan"],
            },
            formula1: { type: "string" },
            formula2: { type: "string" },
            listValues: {
              type: "array",
              items: { type: "string", minLength: 1 },
              minItems: 1,
            },
            allowBlank: { type: "boolean" },
          },
          required: ["type"],
          additionalProperties: false,
        },
      },
      required: ["sheetName", "range", "rule"],
    },
  },
  {
    name: "dataValidation.clear",
    description: "清除区域数据验证",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
    },
  },
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
  ...CHART_IMAGE_TOOL_DEFINITIONS,
  ...STRUCTURE_TOOL_DEFINITIONS,
  ...OBJECT_UPDATE_TOOL_DEFINITIONS,
  ...DISPLAY_TOOL_DEFINITIONS,
  ...FREEZE_TOOL_DEFINITIONS,
  ...PAGE_LAYOUT_TOOL_DEFINITIONS,
  ...SHAPE_TOOL_DEFINITIONS,
  ...TABLE_UNLIST_TOOL_DEFINITIONS,
];
