import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const SHEET_NAMES: JsonSchema = {
  type: "array",
  maxItems: 500,
  items: NON_EMPTY_STRING,
};
const PAGE_BREAKS: JsonSchema = {
  type: "array",
  maxItems: 1_000,
  items: NON_EMPTY_STRING,
};
const SIDE_VALUES = strictNestedObject({
  left: { type: "string" },
  center: { type: "string" },
  right: { type: "string" },
});
const MARGINS = strictNestedObject({
  top: { type: "number", minimum: 0 },
  bottom: { type: "number", minimum: 0 },
  left: { type: "number", minimum: 0 },
  right: { type: "number", minimum: 0 },
  header: { type: "number", minimum: 0 },
  footer: { type: "number", minimum: 0 },
});

const PRINT_TARGET_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["excel", "wps"] },
  actionTimeoutMs: ACTION_TIMEOUT,
  sheetName: NON_EMPTY_STRING,
  sheetNames: SHEET_NAMES,
};

export const EXCEL_PRINT_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "excel",
    operation: "inspectPrintSettings",
    schema: strictNestedObject(PRINT_TARGET_PROPERTIES),
  },
  {
    app: "excel",
    operation: "configurePrint",
    schema: strictNestedObject({
      ...PRINT_TARGET_PROPERTIES,
      orientation: { type: "string", enum: ["portrait", "landscape"] },
      paperSize: {
        type: "string",
        enum: ["a3", "A3", "a4", "A4", "a5", "A5", "letter", "Letter", "legal", "Legal"],
      },
      printArea: NON_EMPTY_STRING,
      repeatRows: NON_EMPTY_STRING,
      repeatColumns: NON_EMPTY_STRING,
      margins: MARGINS,
      marginUnit: { type: "string", enum: ["centimeters", "inches", "points"] },
      fitToOnePageWide: { type: "boolean" },
      fitToOnePageTall: { type: "boolean" },
      scale: { type: "integer", minimum: 10, maximum: 400 },
      fitToPagesWide: { type: "integer", minimum: 1, maximum: 32_767 },
      fitToPagesTall: { type: "integer", minimum: 1, maximum: 32_767 },
      marginLeft: { type: "number", minimum: 0 },
      marginRight: { type: "number", minimum: 0 },
      marginTop: { type: "number", minimum: 0 },
      marginBottom: { type: "number", minimum: 0 },
      centerHorizontally: { type: "boolean" },
      centerVertically: { type: "boolean" },
      printGridlines: { type: "boolean" },
      printHeadings: { type: "boolean" },
      blackAndWhite: { type: "boolean" },
      draft: { type: "boolean" },
      pageOrder: { type: "string", enum: ["downThenOver", "overThenDown"] },
      firstPageNumber: { type: "integer", minimum: 1 },
      header: { type: "string" },
      footer: { type: "string" },
      headers: SIDE_VALUES,
      footers: SIDE_VALUES,
      clearPageBreaks: { type: "boolean" },
      horizontalPageBreaks: PAGE_BREAKS,
      verticalPageBreaks: PAGE_BREAKS,
    }),
  },
];

function strictNestedObject(properties: Record<string, JsonSchema>): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
  };
}
