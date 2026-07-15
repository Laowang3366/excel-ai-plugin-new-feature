import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};

export const OFFICE_EXPORT_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "excel",
    operation: "exportPdf",
    schema: strictObject({
      host: { type: "string", enum: ["excel", "wps"] },
      actionTimeoutMs: ACTION_TIMEOUT,
      scope: { type: "string", enum: ["workbook", "sheet"] },
    }),
  },
  {
    app: "word",
    operation: "exportPdf",
    schema: strictObject({
      host: { type: "string", enum: ["word", "wps"] },
      actionTimeoutMs: ACTION_TIMEOUT,
    }),
  },
  {
    app: "excel",
    operation: "exportSheetsToPdf",
    schema: strictObject({
      host: { type: "string", enum: ["excel", "wps"] },
      actionTimeoutMs: ACTION_TIMEOUT,
      sheetNames: {
        type: "array",
        maxItems: 500,
        items: NON_EMPTY_STRING,
      },
      mode: { type: "string", enum: ["combined", "separate"] },
      outputDirectory: NON_EMPTY_STRING,
      overwrite: { type: "boolean" },
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
