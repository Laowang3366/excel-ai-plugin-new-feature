import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const COLOR: JsonSchema = { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" };
const BASE_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["excel", "wps"] },
  actionTimeoutMs: ACTION_TIMEOUT,
};
const POSITION_PROPERTIES: Record<string, JsonSchema> = {
  left: { type: "number" },
  top: { type: "number" },
  width: { type: "number", minimum: 1 },
  height: { type: "number", minimum: 1 },
};

export const EXCEL_OBJECT_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "excel",
    operation: "inspectWorkbookObjects",
    schema: strictObject(BASE_PROPERTIES),
  },
  workbookObjectSchema(
    "worksheet",
    {
      command: {
        type: "string",
        enum: [
          "add",
          "delete",
          "rename",
          "copy",
          "move",
          "hide",
          "veryHide",
          "show",
          "protect",
          "unprotect",
          "update",
        ],
      },
      sheetName: NON_EMPTY_STRING,
      name: NON_EMPTY_STRING,
      newName: NON_EMPTY_STRING,
      position: { type: "integer", minimum: 1, maximum: 1_048_576 },
      password: { type: "string", maxLength: 255 },
      visible: { type: "boolean" },
      tabColor: COLOR,
    },
    ["command"],
  ),
  workbookObjectSchema(
    "name",
    {
      command: { type: "string", enum: ["create", "update", "delete"] },
      name: NON_EMPTY_STRING,
      refersTo: NON_EMPTY_STRING,
      newName: NON_EMPTY_STRING,
      visible: { type: "boolean" },
    },
    ["command", "name"],
  ),
  workbookObjectSchema(
    "table",
    {
      command: { type: "string", enum: ["create", "update", "delete"] },
      name: NON_EMPTY_STRING,
      newName: NON_EMPTY_STRING,
      style: NON_EMPTY_STRING,
      showTotals: { type: "boolean" },
      unlist: { type: "boolean" },
    },
    ["command"],
  ),
  ...(["chart", "shape"] as const).map((objectType) =>
    workbookObjectSchema(
      objectType,
      {
        command: { type: "string", enum: ["update", "delete"] },
        name: NON_EMPTY_STRING,
        index: { type: "integer", minimum: 1 },
        newName: NON_EMPTY_STRING,
        ...POSITION_PROPERTIES,
      },
      ["command"],
    ),
  ),
  workbookObjectSchema(
    "connection",
    {
      command: { type: "string", enum: ["refresh", "delete"] },
      name: NON_EMPTY_STRING,
    },
    ["command", "name"],
  ),
  {
    app: "excel",
    operation: "manageWorksheetObjects",
    schema: strictObject({
      ...BASE_PROPERTIES,
      command: { type: "string", enum: ["update", "delete"] },
      name: NON_EMPTY_STRING,
      index: { type: "integer", minimum: 1 },
      newName: NON_EMPTY_STRING,
      ...POSITION_PROPERTIES,
    }),
  },
];

function workbookObjectSchema(
  objectType: string,
  properties: Record<string, JsonSchema>,
  required: string[],
): OfficeOperationParamSchema {
  return {
    app: "excel",
    operation: "manageWorkbookObject",
    required: true,
    schema: strictObject(
      {
        ...BASE_PROPERTIES,
        objectType: { type: "string", const: objectType },
        ...properties,
      },
      ["objectType", ...required],
    ),
  };
}

function strictObject(properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
    ...(required.length > 0 ? { required } : {}),
  };
}
