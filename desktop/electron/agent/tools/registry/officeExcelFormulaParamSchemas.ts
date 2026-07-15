import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const SCOPE: JsonSchema = {
  type: "string",
  enum: ["workbook", "sheet", "target"],
};
const BASE_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["excel", "wps"] },
  actionTimeoutMs: ACTION_TIMEOUT,
};
const SCOPED_PROPERTIES: Record<string, JsonSchema> = {
  ...BASE_PROPERTIES,
  scope: SCOPE,
};
const REPLACEMENT = strictObject(
  {
    find: NON_EMPTY_STRING,
    replace: { type: "string" },
  },
  ["find"],
);

export const EXCEL_FORMULA_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  ...(["traceFormulaDependencies", "inspectFormulaDependencies"] as const).map((operation) => ({
    app: "excel" as const,
    operation,
    schema: strictObject(SCOPED_PROPERTIES),
  })),
  {
    app: "excel",
    operation: "repairFormulaReferences",
    required: true,
    schema: strictObject(
      {
        ...SCOPED_PROPERTIES,
        applyAllMappings: { type: "boolean" },
        replacements: {
          type: "array",
          minItems: 1,
          maxItems: 1_000,
          items: REPLACEMENT,
        },
      },
      ["replacements"],
    ),
  },
  {
    app: "excel",
    operation: "convertFormulasToValues",
    schema: strictObject({
      ...SCOPED_PROPERTIES,
      createBackup: { type: "boolean" },
      backupId: NON_EMPTY_STRING,
    }),
  },
  {
    app: "excel",
    operation: "inspectFormulaBackups",
    schema: strictObject(BASE_PROPERTIES),
  },
  {
    app: "excel",
    operation: "restoreFormulas",
    schema: strictObject({
      ...BASE_PROPERTIES,
      backupId: NON_EMPTY_STRING,
      removeAfterRestore: { type: "boolean" },
    }),
  },
  {
    app: "excel",
    operation: "inspectFormulaProtection",
    schema: strictObject(SCOPED_PROPERTIES),
  },
  {
    app: "excel",
    operation: "manageFormulaProtection",
    schema: strictObject({
      ...SCOPED_PROPERTIES,
      command: { type: "string", enum: ["lock", "unlock"] },
      password: { type: "string", maxLength: 255 },
      unlockInputs: { type: "boolean" },
      protectSheet: { type: "boolean" },
    }),
  },
];

function strictObject(properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
    ...(required.length > 0 ? { required } : {}),
  };
}
