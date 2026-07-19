import type { ToolDefinition } from "./types";

export const PAGE_LAYOUT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "sheet.pageLayout.get",
    description:
      "读取工作表页面布局/打印设置（orientation/margins/paperSize/fitToPages/zoomScale/draft/pageOrder/firstPageNumber/printArea 等；ExcelApi 1.9；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.pageLayout.set",
    description:
      "设置页面布局。可选 orientation/center*/print*/blackAndWhite/draft/pageOrder/firstPageNumber/margins/zoomScale/paperSize/fitToPagesWide/fitToPagesTall/printArea/printTitle*；≥1 字段；fit 与 zoomScale 互斥；print* 仅非空（clear 未承诺）；firstPageNumber 仅有限整数≥1。ExcelApi 1.9；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        orientation: { type: "string", enum: ["portrait", "landscape"] },
        centerHorizontally: { type: "boolean" },
        centerVertically: { type: "boolean" },
        printGridlines: { type: "boolean" },
        printHeadings: { type: "boolean" },
        blackAndWhite: { type: "boolean" },
        draft: { type: "boolean" },
        pageOrder: { type: "string", enum: ["downThenOver", "overThenDown"] },
        firstPageNumber: { type: "integer", minimum: 1 },
        margins: {
          type: "object",
          properties: {
            top: { type: "number", minimum: 0 },
            bottom: { type: "number", minimum: 0 },
            left: { type: "number", minimum: 0 },
            right: { type: "number", minimum: 0 },
            header: { type: "number", minimum: 0 },
            footer: { type: "number", minimum: 0 },
          },
          additionalProperties: false,
        },
        zoomScale: { type: "number" },
        paperSize: {
          type: "string",
          enum: ["a3", "a4", "a5", "letter", "legal"],
        },
        fitToPagesWide: { type: "integer", minimum: 1, maximum: 32767 },
        fitToPagesTall: { type: "integer", minimum: 1, maximum: 32767 },
        printArea: { type: "string", minLength: 1 },
        printTitleRows: { type: "string", minLength: 1 },
        printTitleColumns: { type: "string", minLength: 1 },
      },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
];
