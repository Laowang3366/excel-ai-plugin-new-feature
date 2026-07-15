import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const BOUNDED_MATCH_STRING: JsonSchema = { type: "string", minLength: 1, maxLength: 512 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const WORD_BASE_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["word", "wps"] },
  instanceId: NON_EMPTY_STRING,
  actionTimeoutMs: ACTION_TIMEOUT,
};
const VALIDATION_PROPERTIES: Record<string, JsonSchema> = {
  containsText: {
    oneOf: [{ type: "string" }, { type: "array", maxItems: 256, items: { type: "string" } }],
  },
  countPath: NON_EMPTY_STRING,
  expectedCount: { type: "integer", minimum: 0 },
  minCount: { type: "integer", minimum: 0 },
  outputExists: { type: "boolean" },
};
const MARGINS: JsonSchema = strictObject({
  top: { type: "number", minimum: 0, maximum: 100 },
  bottom: { type: "number", minimum: 0, maximum: 100 },
  left: { type: "number", minimum: 0, maximum: 100 },
  right: { type: "number", minimum: 0, maximum: 100 },
});
const HEADER_FOOTER: JsonSchema = strictObject({
  header: { type: "string", maxLength: 32_767 },
  footer: { type: "string", maxLength: 32_767 },
});
const FORMAT_PROPERTIES: Record<string, JsonSchema> = {
  ...WORD_BASE_PROPERTIES,
  fontName: { type: "string", minLength: 1, maxLength: 255 },
  fontSize: { type: "number", minimum: 1, maximum: 200 },
  headerColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
  margins: MARGINS,
  headerFooter: HEADER_FOOTER,
  pageNumbers: { type: "boolean" },
  toc: { type: "string", enum: ["create", "update"] },
  position: { type: "string", enum: ["start", "end"] },
  upperHeadingLevel: { type: "integer", minimum: 1, maximum: 9 },
  lowerHeadingLevel: { type: "integer", minimum: 1, maximum: 9 },
};

export const WORD_FORMATTING_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "word",
    operation: "inspectDocumentFormatting",
    schema: strictObject({ ...WORD_BASE_PROPERTIES, ...VALIDATION_PROPERTIES }),
  },
  formatSchema({ autoDetectHeadings: { type: "boolean", const: true } }, ["autoDetectHeadings"]),
  formatSchema(
    {
      autoDetectHeadings: { type: "boolean", const: false },
      startsWith: BOUNDED_MATCH_STRING,
      level: { type: "integer", minimum: 1, maximum: 9 },
    },
    ["autoDetectHeadings", "startsWith"],
  ),
  formatSchema(
    {
      autoDetectHeadings: { type: "boolean", const: false },
      pattern: BOUNDED_MATCH_STRING,
      level: { type: "integer", minimum: 1, maximum: 9 },
    },
    ["autoDetectHeadings", "pattern"],
  ),
  formatSchema(
    {
      autoDetectHeadings: { type: "boolean", const: false },
      startsWith: BOUNDED_MATCH_STRING,
      pattern: BOUNDED_MATCH_STRING,
      level: { type: "integer", minimum: 1, maximum: 9 },
    },
    ["autoDetectHeadings", "startsWith", "pattern"],
  ),
];

function formatSchema(
  headingProperties: Record<string, JsonSchema>,
  required: string[],
): OfficeOperationParamSchema {
  return {
    app: "word",
    operation: "formatLongDocument",
    required: true,
    schema: strictObject({ ...FORMAT_PROPERTIES, ...headingProperties }, required),
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
