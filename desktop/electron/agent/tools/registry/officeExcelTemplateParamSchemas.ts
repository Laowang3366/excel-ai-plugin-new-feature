import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const BASE_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["excel", "wps"] },
  actionTimeoutMs: ACTION_TIMEOUT,
};

export const EXCEL_TEMPLATE_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  ...(["captureWorkbookTemplate", "inspectWorkbookFormatting"] as const).map((operation) => ({
    app: "excel" as const,
    operation,
    schema: strictObject(BASE_PROPERTIES),
  })),
  {
    app: "excel",
    operation: "applyWorkbookTemplate",
    schema: strictObject({
      ...BASE_PROPERTIES,
      preset: {
        type: "string",
        enum: ["professional", "financial", "dashboard", "minimal"],
      },
      sheetNames: {
        type: "array",
        maxItems: 500,
        items: { type: "string", minLength: 1 },
      },
      allSheets: { type: "boolean" },
      fontName: { type: "string", minLength: 1, maxLength: 255 },
      fontSize: { type: "number", minimum: 1, maximum: 409 },
      autoFit: { type: "boolean" },
      showGridlines: { type: "boolean" },
      freezeRows: { type: "integer", minimum: 0, maximum: 1_048_576 },
    }),
  },
];

function strictObject(properties: Record<string, JsonSchema>): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
  };
}
