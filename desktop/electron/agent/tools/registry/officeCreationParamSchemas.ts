import type {
  JsonSchema,
  OfficeOperationParamSchema,
  OfficeSchemaApp,
} from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const APP_HOSTS: Record<OfficeSchemaApp, JsonSchema> = {
  excel: { type: "string", enum: ["excel", "wps"] },
  word: { type: "string", enum: ["word", "wps"] },
  presentation: { type: "string", enum: ["powerpoint", "wps"] },
};

function strictAppObject(app: OfficeSchemaApp, properties: Record<string, JsonSchema>): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties: { ...properties, host: APP_HOSTS[app] },
  };
}

export const OFFICE_CREATION_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "excel",
    operation: "createWorkbook",
    schema: strictAppObject("excel", {
      sheetNames: { type: "array", minItems: 1, maxItems: 1_000, items: NON_EMPTY_STRING },
      values: {
        type: "array",
        maxItems: 20_000,
        items: { type: "array", maxItems: 20_000, items: {} },
      },
      startCell: NON_EMPTY_STRING,
    }),
  },
  {
    app: "word",
    operation: "createDocument",
    schema: strictAppObject("word", {
      title: { type: "string" },
      paragraphs: {
        oneOf: [{ type: "string" }, { type: "array", maxItems: 20_000, items: { type: "string" } }],
      },
      text: { type: "string" },
      body: { type: "string" },
    }),
  },
  {
    app: "presentation",
    operation: "createPresentation",
    schema: strictAppObject("presentation", {
      title: { type: "string" },
      subtitle: { type: "string" },
    }),
  },
];
