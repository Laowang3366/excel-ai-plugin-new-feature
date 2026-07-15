import { OFFICE_CAPABILITIES } from "../officeCore/capabilities";
import { SAFE_ACTION_OPERATIONS } from "../officeCore/operationPolicy";
import {
  MODELED_OFFICE_PARAM_SCHEMAS,
  type OfficeOperationParamSchema,
} from "./officeOperationParamSchemas";

type JsonSchema = Record<string, any>;

export const POWER_QUERY_OPERATIONS = new Set([
  "createPowerQuery",
  "managePowerQuery",
  "inspectPowerQueries",
]);

export const PIVOT_OPERATIONS = new Set(["createPivotTable", "refreshPivotTables", "addSlicer"]);

export const APPLY_OPERATIONS = uniqueSorted(
  OFFICE_CAPABILITIES.filter((capability) => capability.writesFile).map(
    (capability) => capability.operation,
  ),
);

export const INSPECT_OPERATIONS = uniqueSorted(Array.from(SAFE_ACTION_OPERATIONS));

export const WORKFLOW_OPERATIONS = uniqueSorted([...APPLY_OPERATIONS, ...INSPECT_OPERATIONS]);

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ADVANCED_ETL: JsonSchema = { type: "string", enum: ["refreshable-etl"] };
const INTERACTIVE_PIVOT: JsonSchema = { type: "string", enum: ["interactive-pivot"] };

const POWER_QUERY_PROPERTIES: Record<string, JsonSchema> = {
  advancedIntent: ADVANCED_ETL,
  sourceKind: { type: "string", enum: ["external", "multi-source"] },
  command: {
    type: "string",
    enum: [
      "create",
      "update",
      "upsert",
      "duplicate",
      "rename",
      "load",
      "refresh",
      "unload",
      "delete",
    ],
  },
  name: NON_EMPTY_STRING,
  newName: NON_EMPTY_STRING,
  mFormula: NON_EMPTY_STRING,
  description: { type: "string" },
  loadMode: { type: "string", enum: ["worksheet", "dataModel", "connectionOnly"] },
  destination: NON_EMPTY_STRING,
  tableName: NON_EMPTY_STRING,
  refresh: { type: "boolean" },
  clearOutput: { type: "boolean" },
};

const PIVOT_FIELD: JsonSchema = {
  oneOf: [
    NON_EMPTY_STRING,
    {
      type: "object",
      additionalProperties: false,
      properties: {
        name: NON_EMPTY_STRING,
        function: {
          type: "string",
          enum: ["sum", "count", "average", "max", "min"],
        },
        caption: { type: "string" },
      },
      required: ["name"],
    },
  ],
};

const PIVOT_FIELDS: JsonSchema = {
  type: "array",
  maxItems: 64,
  items: PIVOT_FIELD,
};

const ADVANCED_PARAM_SCHEMAS: Record<string, { schema: JsonSchema; required: boolean }> = {
  createPowerQuery: {
    required: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: POWER_QUERY_PROPERTIES,
      required: ["advancedIntent", "sourceKind", "name", "mFormula", "loadMode"],
    },
  },
  managePowerQuery: {
    required: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: POWER_QUERY_PROPERTIES,
      required: ["advancedIntent", "name"],
    },
  },
  inspectPowerQueries: {
    required: false,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        name: NON_EMPTY_STRING,
        containsText: {
          oneOf: [{ type: "string" }, { type: "array", maxItems: 256, items: { type: "string" } }],
        },
        countPath: NON_EMPTY_STRING,
        expectedCount: { type: "integer", minimum: 0 },
        minCount: { type: "integer", minimum: 0 },
        outputExists: { type: "boolean" },
      },
    },
  },
  createPivotTable: {
    required: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        advancedIntent: INTERACTIVE_PIVOT,
        name: NON_EMPTY_STRING,
        destination: NON_EMPTY_STRING,
        rowFields: PIVOT_FIELDS,
        columnFields: PIVOT_FIELDS,
        filterFields: PIVOT_FIELDS,
        dataFields: PIVOT_FIELDS,
      },
      required: ["advancedIntent"],
    },
  },
  refreshPivotTables: {
    required: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        advancedIntent: INTERACTIVE_PIVOT,
        refreshConnections: { type: "boolean" },
      },
      required: ["advancedIntent"],
    },
  },
  addSlicer: {
    required: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        advancedIntent: INTERACTIVE_PIVOT,
        pivotName: NON_EMPTY_STRING,
        field: NON_EMPTY_STRING,
        name: NON_EMPTY_STRING,
        caption: { type: "string" },
        top: { type: "number" },
        left: { type: "number" },
      },
      required: ["advancedIntent", "pivotName", "field"],
    },
  },
};

const OPERATION_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  ...Object.entries(ADVANCED_PARAM_SCHEMAS).map(([operation, definition]) => ({
    app: "excel" as const,
    operation,
    ...definition,
  })),
  ...MODELED_OFFICE_PARAM_SCHEMAS,
];

export function withOfficeOperationDiscriminator(
  schema: JsonSchema,
  operations: string[],
): JsonSchema {
  const baseSchema = { ...schema };
  delete baseSchema.oneOf;
  const properties = asSchema(baseSchema.properties);
  const required = Array.isArray(baseSchema.required) ? baseSchema.required : [];
  const variants = operations.flatMap((operation) =>
    OPERATION_PARAM_SCHEMAS.filter((definition) => definition.operation === operation).map(
      (definition) => operationVariant(baseSchema, properties, required, definition),
    ),
  );
  const ordinaryOperations = operations.filter(
    (operation) =>
      !OPERATION_PARAM_SCHEMAS.some((definition) => definition.operation === operation),
  );
  if (ordinaryOperations.length > 0) {
    variants.push({
      ...baseSchema,
      properties: {
        ...properties,
        operation: { ...asSchema(properties.operation), enum: ordinaryOperations },
      },
    });
  }
  return { ...baseSchema, oneOf: variants };
}

function operationVariant(
  schema: JsonSchema,
  properties: JsonSchema,
  required: unknown[],
  definition: OfficeOperationParamSchema,
): JsonSchema {
  return {
    ...schema,
    properties: {
      ...properties,
      app: definition.app
        ? { ...asSchema(properties.app), const: definition.app }
        : asSchema(properties.app),
      operation: { ...asSchema(properties.operation), const: definition.operation },
      params: definition.schema,
    },
    required: definition.required ? Array.from(new Set([...required, "params"])) : required,
  };
}

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values)).sort();
}

function asSchema(value: unknown): JsonSchema {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as JsonSchema) : {};
}
