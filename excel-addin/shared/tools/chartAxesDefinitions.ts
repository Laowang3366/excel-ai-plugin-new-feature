import type { ToolDefinition } from "./types";

const unitSchema = {
  anyOf: [{ type: "number", minimum: 0 }, { type: "string", const: "" }],
};

export const CHART_AXES_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.axes.update",
    description:
      "更新图表坐标轴：kind category|value；group primary|secondary（默认 primary）；title/minimum/maximum/majorUnit|minorUnit(数值或空串=自动)/numberFormat/reverse；displayUnit/customDisplayUnit/scaleType/logBase/showDisplayUnitLabel/majorTickMark/minorTickMark/tickLabelPosition（ExcelApi 1.7）；position/positionAt（ExcelApi 1.8，position=custom 需 positionAt）；linkNumberFormat（ExcelApi 1.9）；majorGridlinesVisible/minorGridlinesVisible（ExcelApi 1.1）。≥1 更新字段；title 空串清除。写后宿主回读。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        kind: { type: "string", enum: ["category", "value"] },
        group: { type: "string", enum: ["primary", "secondary"] },
        title: { type: "string" },
        minimum: { type: "number" },
        maximum: { type: "number" },
        majorUnit: unitSchema,
        minorUnit: { anyOf: [{ type: "number" }, { type: "string", const: "" }] },
        numberFormat: { type: "string", minLength: 1 },
        reverse: { type: "boolean" },
        displayUnit: {
          type: "string",
          enum: [
            "none",
            "hundreds",
            "thousands",
            "tenThousands",
            "hundredThousands",
            "millions",
            "tenMillions",
            "hundredMillions",
            "billions",
            "trillions",
            "custom",
          ],
        },
        customDisplayUnit: { type: "number" },
        scaleType: { type: "string", enum: ["linear", "logarithmic"] },
        logBase: { type: "number" },
        showDisplayUnitLabel: { type: "boolean" },
        majorGridlinesVisible: { type: "boolean" },
        minorGridlinesVisible: { type: "boolean" },
        majorTickMark: { type: "string", enum: ["none", "cross", "inside", "outside"] },
        minorTickMark: { type: "string", enum: ["none", "cross", "inside", "outside"] },
        tickLabelPosition: {
          type: "string",
          enum: ["nextToAxis", "high", "low", "none"],
        },
        position: {
          type: "string",
          enum: ["automatic", "maximum", "minimum", "custom"],
        },
        positionAt: { type: "number" },
        linkNumberFormat: { type: "boolean" },
      },
      required: ["sheetName", "chartName", "kind"],
      additionalProperties: false,
    },
  },
];
