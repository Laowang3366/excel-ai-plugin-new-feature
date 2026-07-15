import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const EXCEL_HOST: JsonSchema = { type: "string", enum: ["excel", "wps"] };
const WORD_HOST: JsonSchema = { type: "string", enum: ["word", "wps"] };
const PRESENTATION_HOST: JsonSchema = { type: "string", enum: ["powerpoint", "wps"] };

const VALIDATION_PROPERTIES: Record<string, JsonSchema> = {
  containsText: {
    oneOf: [{ type: "string" }, { type: "array", maxItems: 256, items: { type: "string" } }],
  },
  countPath: NON_EMPTY_STRING,
  expectedCount: { type: "integer", minimum: 0 },
  minCount: { type: "integer", minimum: 0 },
  outputExists: { type: "boolean" },
};

const SOURCE_ROUTING: Record<string, JsonSchema> = {
  sourceHost: EXCEL_HOST,
  sourceInstanceId: NON_EMPTY_STRING,
};

const PRESENTATION_GEOMETRY: Record<string, JsonSchema> = {
  left: { type: "number", minimum: 0 },
  top: { type: "number", minimum: 0 },
  width: { type: "number", minimum: 1 },
};

const REPORT_SECTION_BASE: Record<string, JsonSchema> = {
  linkId: NON_EMPTY_STRING,
  sheetName: NON_EMPTY_STRING,
  range: NON_EMPTY_STRING,
  title: { type: "string" },
  ...PRESENTATION_GEOMETRY,
};

export const CROSS_OFFICE_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "excel",
    operation: "exportRangeToWord",
    required: true,
    schema: exportSchema("word"),
  },
  {
    app: "excel",
    operation: "exportRangeToPresentation",
    required: true,
    schema: exportSchema("presentation"),
  },
  {
    app: "excel",
    operation: "buildReportPackage",
    required: true,
    schema: {
      oneOf: [reportSchema(false), reportSchema(true)],
    },
  },
  ...linkedContentSchemas("word"),
  ...linkedContentSchemas("presentation"),
];

function exportSchema(target: "word" | "presentation"): JsonSchema {
  return {
    oneOf: [
      ...[false, true].flatMap((updateExisting) =>
        [false, true].flatMap((linked) => [
          exportVariant(target, "range", updateExisting, linked),
          exportVariant(target, "chart", updateExisting, linked),
        ]),
      ),
    ],
  };
}

function exportVariant(
  target: "word" | "presentation",
  sourceType: "range" | "chart",
  updateExisting: boolean,
  linked: boolean,
): JsonSchema {
  const targetRouting: Record<string, JsonSchema> =
    target === "word"
      ? { wordHost: WORD_HOST, instanceId: NON_EMPTY_STRING }
      : { presentationHost: PRESENTATION_HOST, instanceId: NON_EMPTY_STRING };
  const targetProperties: Record<string, JsonSchema> =
    target === "word"
      ? linked
        ? {}
        : { asPicture: { type: "boolean" } as JsonSchema }
      : PRESENTATION_GEOMETRY;
  return strictObject(
    {
      ...SOURCE_ROUTING,
      ...targetRouting,
      ...targetProperties,
      actionTimeoutMs: ACTION_TIMEOUT,
      linked: { type: "boolean", const: linked },
      linkId: NON_EMPTY_STRING,
      updateExisting: { type: "boolean", const: updateExisting },
      ...(updateExisting
        ? { allowMissingManaged: { type: "boolean" } as JsonSchema }
        : { overwrite: { type: "boolean" } as JsonSchema }),
      title: { type: "string" },
      sourceType: { type: "string", const: sourceType },
      ...(sourceType === "chart" ? { chartName: NON_EMPTY_STRING } : {}),
    },
    [
      "linked",
      ...(updateExisting ? ["updateExisting", "linkId"] : []),
      ...(sourceType === "chart" ? ["sourceType", "chartName"] : []),
    ],
  );
}

function reportSchema(updateExisting: boolean): JsonSchema {
  return strictObject(
    {
      ...SOURCE_ROUTING,
      wordHost: WORD_HOST,
      wordInstanceId: NON_EMPTY_STRING,
      presentationHost: PRESENTATION_HOST,
      presentationInstanceId: NON_EMPTY_STRING,
      actionTimeoutMs: ACTION_TIMEOUT,
      outputDirectory: NON_EMPTY_STRING,
      baseName: NON_EMPTY_STRING,
      wordOutputPath: NON_EMPTY_STRING,
      presentationOutputPath: NON_EMPTY_STRING,
      linked: { type: "boolean", const: true },
      updateExisting: { type: "boolean", const: updateExisting },
      ...(updateExisting ? {} : { overwrite: { type: "boolean" } as JsonSchema }),
      sections: {
        type: "array",
        minItems: 1,
        maxItems: 100,
        items: {
          oneOf: [reportSection("range", updateExisting), reportSection("chart", updateExisting)],
        },
      },
    },
    ["linked", "sections", ...(updateExisting ? ["updateExisting"] : [])],
  );
}

function reportSection(sourceType: "range" | "chart", updateExisting: boolean): JsonSchema {
  return strictObject(
    {
      ...REPORT_SECTION_BASE,
      sourceType: { type: "string", const: sourceType },
      ...(sourceType === "chart" ? { chartName: NON_EMPTY_STRING } : {}),
    },
    [
      "range",
      ...(updateExisting ? ["linkId"] : []),
      ...(sourceType === "chart" ? ["sourceType", "chartName"] : []),
    ],
  );
}

function linkedContentSchemas(app: "word" | "presentation"): OfficeOperationParamSchema[] {
  const baseProperties: Record<string, JsonSchema> = {
    host: app === "word" ? WORD_HOST : PRESENTATION_HOST,
    instanceId: NON_EMPTY_STRING,
    ...SOURCE_ROUTING,
    actionTimeoutMs: ACTION_TIMEOUT,
  };
  return [
    {
      app,
      operation: "inspectLinkedOfficeContent",
      schema: strictObject({
        ...baseProperties,
        linkId: NON_EMPTY_STRING,
        ...VALIDATION_PROPERTIES,
      }),
    },
    {
      app,
      operation: "refreshLinkedOfficeContent",
      schema: strictObject({ ...baseProperties, linkId: NON_EMPTY_STRING }),
    },
    {
      app,
      operation: "relinkLinkedOfficeContent",
      required: true,
      schema: strictObject(
        { ...baseProperties, linkId: NON_EMPTY_STRING, sourcePath: NON_EMPTY_STRING },
        ["linkId", "sourcePath"],
      ),
    },
  ];
}

function strictObject(properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
    ...(required.length > 0 ? { required } : {}),
  };
}
