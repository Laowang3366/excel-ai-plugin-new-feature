import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const OPEN_XML_SPREADSHEET_PATH: JsonSchema = {
  type: "string",
  minLength: 1,
  pattern: "\\.(?:[Xx][Ll][Ss][XxMm]|[Xx][Ll][Tt][XxMm])$",
};
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
const CONTROL_TYPES = [
  "richtext",
  "text",
  "picture",
  "combobox",
  "dropdown",
  "dropdownlist",
  "date",
  "checkbox",
] as const;
const LIST_CONTROL_TYPES = new Set<string>(["combobox", "dropdown", "dropdownlist"]);
const CONTROL_ENTRY: JsonSchema = strictObject(
  { text: NON_EMPTY_STRING, value: { type: "string" } },
  ["text"],
);
const CONTROL_SPEC: JsonSchema = {
  oneOf: [controlSpecSchema(), ...CONTROL_TYPES.map((type) => controlSpecSchema(type))],
};
const CONTROL_VALUE: JsonSchema = {
  oneOf: [
    { type: "string" },
    { type: "number" },
    { type: "boolean" },
    strictObject(
      {
        value: {
          oneOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }],
        },
        dateFormat: { type: "string", minLength: 1, maxLength: 64 },
      },
      ["value"],
    ),
  ],
};
const CONDITION: JsonSchema = strictObject(
  {
    placeholder: NON_EMPTY_STRING,
    field: NON_EMPTY_STRING,
    operator: { type: "string", enum: ["eq", "ne", "contains"] },
    value: { type: "string" },
    trueText: { type: "string" },
    falseText: { type: "string" },
  },
  ["placeholder", "field"],
);
const IMAGE_FIELD: JsonSchema = strictObject(
  {
    placeholder: NON_EMPTY_STRING,
    field: NON_EMPTY_STRING,
    width: { type: "number", minimum: 1, maximum: 10_000 },
  },
  ["placeholder", "field"],
);
const MERGE_PROPERTIES: Record<string, JsonSchema> = {
  ...WORD_BASE_PROPERTIES,
  dataSourcePath: OPEN_XML_SPREADSHEET_PATH,
  outputFormat: { type: "string", enum: ["docx", "pdf", "both"] },
  conditions: { type: "array", maxItems: 256, items: CONDITION },
  imageFields: { type: "array", maxItems: 256, items: IMAGE_FIELD },
};

export const WORD_TEMPLATE_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "word",
    operation: "inspectContentControls",
    schema: strictObject({ ...WORD_BASE_PROPERTIES, ...VALIDATION_PROPERTIES }),
  },
  {
    app: "word",
    operation: "prepareMailMergeTemplate",
    required: true,
    schema: strictObject(
      {
        ...WORD_BASE_PROPERTIES,
        fields: {
          type: "array",
          minItems: 1,
          maxItems: 500,
          items: {
            oneOf: [NON_EMPTY_STRING, strictObject({ name: NON_EMPTY_STRING }, ["name"])],
          },
        },
      },
      ["fields"],
    ),
  },
  {
    app: "word",
    operation: "populateContentControls",
    required: true,
    schema: strictObject(
      {
        ...WORD_BASE_PROPERTIES,
        values: {
          type: "object",
          minProperties: 1,
          maxProperties: 512,
          propertyNames: {
            type: "string",
            pattern: "^(?!__proto__$|prototype$|constructor$).{1,255}$",
          },
          additionalProperties: CONTROL_VALUE,
        },
        dateFormat: { type: "string", minLength: 1, maxLength: 64 },
      },
      ["values"],
    ),
  },
  {
    app: "word",
    operation: "mailMerge",
    required: true,
    schema: strictObject(MERGE_PROPERTIES, ["dataSourcePath"]),
  },
  {
    app: "word",
    operation: "batchMailMerge",
    required: true,
    schema: strictObject(
      {
        ...MERGE_PROPERTIES,
        outputDirectory: NON_EMPTY_STRING,
        fileNamePattern: { type: "string", minLength: 1, maxLength: 240 },
      },
      ["dataSourcePath"],
    ),
  },
  ...addControlSchemas(),
  ...deleteControlSchemas(),
  ...updateControlSchemas(),
];

function addControlSchemas(): OfficeOperationParamSchema[] {
  const schemas = [
    manageSchema(
      {
        command: { type: "string", const: "add" },
        ...WORD_BASE_PROPERTIES,
        controls: { type: "array", minItems: 1, maxItems: 100, items: CONTROL_SPEC },
      },
      ["command", "controls"],
    ),
  ];
  for (const type of [undefined, ...CONTROL_TYPES] as const) {
    schemas.push(
      manageSchema(
        {
          command: { type: "string", const: "add" },
          ...WORD_BASE_PROPERTIES,
          ...controlSpecProperties(type),
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 0 },
        },
        type ? ["command", "type"] : ["command"],
      ),
    );
  }
  return schemas;
}

function deleteControlSchemas(): OfficeOperationParamSchema[] {
  return ["id", "title", "tag"].map((selector) =>
    manageSchema(
      {
        command: { type: "string", const: "delete" },
        ...WORD_BASE_PROPERTIES,
        [selector]: NON_EMPTY_STRING,
        deleteContents: { type: "boolean" },
      },
      ["command", selector],
    ),
  );
}

function updateControlSchemas(): OfficeOperationParamSchema[] {
  const updateProperties: Record<string, JsonSchema> = {
    title: NON_EMPTY_STRING,
    tag: NON_EMPTY_STRING,
    lockContents: { type: "boolean" },
    lockControl: { type: "boolean" },
  };
  const keys = Object.keys(updateProperties);
  const schemas: OfficeOperationParamSchema[] = [];
  for (let mask = 1; mask < 1 << keys.length; mask += 1) {
    const selected = keys.filter((_, index) => (mask & (1 << index)) !== 0);
    schemas.push(
      manageSchema(
        {
          command: { type: "string", const: "update" },
          ...WORD_BASE_PROPERTIES,
          id: NON_EMPTY_STRING,
          ...Object.fromEntries(selected.map((key) => [key, updateProperties[key]])),
        },
        ["command", "id", ...selected],
      ),
    );
  }
  for (const selector of ["title", "tag"] as const) {
    for (const locks of [["lockContents"], ["lockControl"], ["lockContents", "lockControl"]]) {
      schemas.push(
        manageSchema(
          {
            command: { type: "string", const: "update" },
            ...WORD_BASE_PROPERTIES,
            [selector]: NON_EMPTY_STRING,
            ...Object.fromEntries(locks.map((key) => [key, updateProperties[key]])),
          },
          ["command", selector, ...locks],
        ),
      );
    }
  }
  return schemas;
}

function controlSpecSchema(type?: (typeof CONTROL_TYPES)[number]): JsonSchema {
  return strictObject(controlSpecProperties(type), type ? ["type"] : []);
}

function controlSpecProperties(type?: (typeof CONTROL_TYPES)[number]): Record<string, JsonSchema> {
  return {
    ...(type ? { type: { type: "string", const: type } } : {}),
    title: { type: "string" },
    tag: { type: "string" },
    placeholder: { type: "string" },
    lockContents: { type: "boolean" },
    lockControl: { type: "boolean" },
    ...(type && LIST_CONTROL_TYPES.has(type)
      ? { entries: { type: "array", maxItems: 256, items: CONTROL_ENTRY } }
      : {}),
  };
}

function manageSchema(
  properties: Record<string, JsonSchema>,
  required: string[],
): OfficeOperationParamSchema {
  return {
    app: "word",
    operation: "manageContentControls",
    required: true,
    schema: strictObject(properties, required),
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
