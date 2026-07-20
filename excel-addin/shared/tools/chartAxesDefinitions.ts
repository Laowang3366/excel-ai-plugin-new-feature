import type { ToolDefinition } from "./types";

export const CHART_AXES_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.axes.update",
    description:
      "更新图表坐标轴：kind category|value；group primary|secondary（默认 primary）；title/minimum/maximum/majorUnit/numberFormat/reverse；displayUnit/customDisplayUnit/scaleType/logBase/showDisplayUnitLabel（ExcelApi 1.7）；majorGridlinesVisible/minorGridlinesVisible（ExcelApi 1.1）。≥1 更新字段；title 空串清除。写后宿主回读。WPS unsupported",
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
        majorUnit: { type: "number", minimum: 0 },
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
      },
      required: ["sheetName", "chartName", "kind"],
      additionalProperties: false,
    },
  },
];
