import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
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
const REFERENCE_TARGET_PROPERTIES: Record<string, JsonSchema> = {
  bookmark: NON_EMPTY_STRING,
  start: { type: "integer", minimum: 0 },
  end: { type: "integer", minimum: 0 },
};
const TRACKED_REPLACE_PROPERTIES: Record<string, JsonSchema> = {
  find: NON_EMPTY_STRING,
  replace: { type: "string" },
  matchCase: { type: "boolean" },
  all: { type: "boolean" },
};

const TRACKED_CHANGE: JsonSchema = {
  oneOf: [
    strictObject(
      {
        command: { type: "string", const: "insert" },
        text: NON_EMPTY_STRING,
        position: {
          oneOf: [
            { type: "string", enum: ["start", "end"] },
            { type: "integer", minimum: 0 },
          ],
        },
      },
      ["command", "text"],
    ),
    strictObject(
      {
        command: { type: "string", const: "replaceBookmark" },
        name: NON_EMPTY_STRING,
        text: { type: "string" },
      },
      ["command", "name", "text"],
    ),
    contentControlChange({ tag: NON_EMPTY_STRING }, ["tag"]),
    contentControlChange({ title: NON_EMPTY_STRING }, ["title"]),
    contentControlChange({ tag: NON_EMPTY_STRING, title: NON_EMPTY_STRING }, ["tag", "title"]),
    strictObject(
      {
        command: { type: "string", const: "replace" },
        ...TRACKED_REPLACE_PROPERTIES,
      },
      ["command", "find"],
    ),
    strictObject(
      {
        command: { type: "string", const: "delete" },
        find: NON_EMPTY_STRING,
        matchCase: { type: "boolean" },
        all: { type: "boolean" },
      },
      ["command", "find"],
    ),
  ],
};

export const WORD_REVIEW_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  ...(["inspectReferences", "inspectRevisions"] as const).map((operation) => ({
    app: "word" as const,
    operation,
    schema: strictObject({ ...WORD_BASE_PROPERTIES, ...VALIDATION_PROPERTIES }),
  })),
  ...referenceCommandSchemas(),
  ...revisionCommandSchemas(),
  compareSchema("comparePath"),
  compareSchema("revisedFilePath"),
  trackedChangesSchema("changes"),
  trackedChangesSchema("edits"),
];

function referenceCommandSchemas(): OfficeOperationParamSchema[] {
  const common = { ...WORD_BASE_PROPERTIES, ...REFERENCE_TARGET_PROPERTIES };
  return [
    ...(["createBookmark", "addBookmark", "deleteBookmark"] as const).map((command) =>
      commandSchema("manageReferences", command, { ...common, name: NON_EMPTY_STRING }, ["name"]),
    ),
    ...(["addFootnote", "addEndnote"] as const).map((command) =>
      commandSchema("manageReferences", command, { ...common, text: NON_EMPTY_STRING }, ["text"]),
    ),
    commandSchema("manageReferences", "addCaption", {
      ...common,
      label: NON_EMPTY_STRING,
      title: { type: "string" },
    }),
    commandSchema(
      "manageReferences",
      "addCrossReference",
      { ...common, referenceType: NON_EMPTY_STRING, item: NON_EMPTY_STRING },
      ["item"],
    ),
    commandSchema("manageReferences", "addTableOfFigures", {
      ...common,
      label: NON_EMPTY_STRING,
    }),
    commandSchema("manageReferences", "updateFields", common),
  ];
}

function revisionCommandSchemas(): OfficeOperationParamSchema[] {
  return [
    ...(["acceptAll", "rejectAll"] as const).map((command) =>
      commandSchema("manageRevisions", command, WORD_BASE_PROPERTIES),
    ),
    ...(["accept", "reject"] as const).map((command) =>
      commandSchema("manageRevisions", command, {
        ...WORD_BASE_PROPERTIES,
        author: NON_EMPTY_STRING,
        revisionType: { type: "integer", minimum: 0, maximum: 64 },
      }),
    ),
    commandSchema(
      "manageRevisions",
      "track",
      { ...WORD_BASE_PROPERTIES, enabled: { type: "boolean" } },
      ["enabled"],
    ),
  ];
}

function commandSchema(
  operation: string,
  command: string,
  properties: Record<string, JsonSchema>,
  required: string[] = [],
): OfficeOperationParamSchema {
  return {
    app: "word",
    operation,
    required: true,
    schema: strictObject({ ...properties, command: { type: "string", const: command } }, [
      "command",
      ...required,
    ]),
  };
}

function compareSchema(
  pathProperty: "comparePath" | "revisedFilePath",
): OfficeOperationParamSchema {
  return {
    app: "word",
    operation: "compareDocuments",
    required: true,
    schema: strictObject({ ...WORD_BASE_PROPERTIES, [pathProperty]: NON_EMPTY_STRING }, [
      pathProperty,
    ]),
  };
}

function trackedChangesSchema(property: "changes" | "edits"): OfficeOperationParamSchema {
  return {
    app: "word",
    operation: "applyTrackedChanges",
    required: true,
    schema: strictObject(
      {
        ...WORD_BASE_PROPERTIES,
        [property]: { type: "array", minItems: 1, maxItems: 1_000, items: TRACKED_CHANGE },
        keepTracking: { type: "boolean" },
        restoreTracking: { type: "boolean" },
      },
      [property],
    ),
  };
}

function contentControlChange(
  selector: Record<string, JsonSchema>,
  requiredSelector: string[],
): JsonSchema {
  return strictObject(
    {
      command: { type: "string", const: "replaceContentControl" },
      ...selector,
      text: { type: "string" },
    },
    ["command", ...requiredSelector, "text"],
  );
}

function strictObject(properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties,
    ...(required.length > 0 ? { required } : {}),
  };
}
