export type OfficeSchemaApp = "excel" | "word" | "presentation";

export type JsonSchema = Record<string, unknown>;

export interface OfficeOperationParamSchema {
  app?: OfficeSchemaApp;
  operation: string;
  schema: JsonSchema;
  required?: boolean;
}

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const COLOR: JsonSchema = { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const CHART_TYPE: JsonSchema = {
  type: "string",
  enum: ["line", "pie", "bar", "area", "scatter", "column"],
};
const TABLE_STYLE: JsonSchema = {
  type: "string",
  enum: ["professional", "compact", "financial"],
};
const APP_HOSTS: Record<OfficeSchemaApp, JsonSchema> = {
  excel: { type: "string", enum: ["excel", "wps"] },
  word: { type: "string", enum: ["word", "wps"] },
  presentation: { type: "string", enum: ["powerpoint", "wps"] },
};

const VALIDATION_PROPERTIES: Record<string, JsonSchema> = {
  containsText: {
    oneOf: [{ type: "string" }, { type: "array", maxItems: 256, items: { type: "string" } }],
  },
  countPath: NON_EMPTY_STRING,
  expectedCount: { type: "integer", minimum: 0 },
  minCount: { type: "integer", minimum: 0 },
  outputExists: { type: "boolean" },
  actionTimeoutMs: ACTION_TIMEOUT,
};

const SLIDE_PROPERTIES: Record<string, JsonSchema> = {
  title: { type: "string" },
  body: { type: "string" },
  bullets: { type: "array", maxItems: 200, items: { type: "string" } },
  layout: {
    type: "string",
    enum: ["title", "titleOnly", "titleAndContent", "blank"],
  },
  index: { type: "integer", minimum: 1 },
};

const POSITION_PROPERTIES: Record<string, JsonSchema> = {
  left: { type: "number" },
  top: { type: "number" },
  width: { type: "number", minimum: 1 },
  height: { type: "number", minimum: 1 },
  name: { type: "string" },
};

function strictObject(properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema {
  return {
    type: "object",
    additionalProperties: false,
    properties: { ...properties, actionTimeoutMs: ACTION_TIMEOUT },
    ...(required.length > 0 ? { required } : {}),
  };
}

function strictAppObject(
  app: OfficeSchemaApp,
  properties: Record<string, JsonSchema>,
  required: string[] = [],
): JsonSchema {
  return strictObject({ ...properties, host: APP_HOSTS[app] }, required);
}

const COMMON_INSPECTION_OPERATIONS = ["inspectFile", "layout", "tables", "listBackups"];

export const MODELED_OFFICE_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  ...COMMON_INSPECTION_OPERATIONS.map((operation) => ({
    operation,
    schema: strictObject(VALIDATION_PROPERTIES),
  })),
  ...(["excel", "word", "presentation"] as const).map((app) => ({
    app,
    operation: "snapshot",
    schema: strictAppObject(app, {}),
  })),
  {
    app: "excel",
    operation: "insertChart",
    schema: strictAppObject("excel", { chartType: CHART_TYPE }),
  },
  {
    app: "excel",
    operation: "applyConditionalFormatting",
    schema: strictAppObject("excel", { formula: { type: "string" }, fillColor: COLOR }),
  },
  {
    app: "excel",
    operation: "setDataValidation",
    schema: strictAppObject("excel", {
      formula: NON_EMPTY_STRING,
      list: NON_EMPTY_STRING,
      values: { type: "array", maxItems: 1_000, items: { type: "string" } },
      type: { type: "string" },
    }),
  },
  {
    app: "excel",
    operation: "styleTable",
    schema: strictAppObject("excel", { style: TABLE_STYLE, headerColor: COLOR }),
  },
  {
    app: "word",
    operation: "applyHeadingStyles",
    schema: strictAppObject("word", {
      startsWith: { type: "string" },
      pattern: { type: "string" },
      level: { type: "integer", minimum: 1, maximum: 9 },
    }),
  },
  {
    app: "word",
    operation: "insertOrUpdateToc",
    schema: strictAppObject("word", {
      position: { type: "string", enum: ["start", "end"] },
      upperHeadingLevel: { type: "integer", minimum: 1, maximum: 9 },
      lowerHeadingLevel: { type: "integer", minimum: 1, maximum: 9 },
    }),
  },
  {
    app: "word",
    operation: "styleTables",
    schema: strictAppObject("word", { style: TABLE_STYLE, headerColor: COLOR }),
  },
  {
    app: "word",
    operation: "setHeaderFooter",
    schema: strictAppObject("word", {
      kind: { type: "string", enum: ["header", "footer"] },
      text: { type: "string" },
    }),
  },
  {
    app: "word",
    operation: "insertOrReplaceImage",
    required: true,
    schema: strictAppObject(
      "word",
      {
        imagePath: NON_EMPTY_STRING,
        bookmark: { type: "string" },
        width: { type: "number", minimum: 1 },
        height: { type: "number", minimum: 1 },
      },
      ["imagePath"],
    ),
  },
  ...(["addSlide", "appendSlide"] as const).map((operation) => ({
    app: "presentation" as const,
    operation,
    schema: strictAppObject("presentation", SLIDE_PROPERTIES),
  })),
  ...(["addSlides", "appendSlides"] as const).map((operation) => ({
    app: "presentation" as const,
    operation,
    required: true,
    schema: strictAppObject(
      "presentation",
      {
        slides: {
          type: "array",
          minItems: 1,
          maxItems: 100,
          items: {
            type: "object",
            additionalProperties: false,
            properties: SLIDE_PROPERTIES,
          },
        },
      },
      ["slides"],
    ),
  })),
  {
    app: "presentation",
    operation: "addSlideContent",
    schema: strictAppObject("presentation", {
      contentType: { type: "string", enum: ["text", "image"] },
      type: { type: "string", enum: ["text", "image"] },
      imagePath: { type: "string" },
      text: { type: "string" },
      ...POSITION_PROPERTIES,
    }),
  },
  {
    app: "presentation",
    operation: "applyTheme",
    schema: strictAppObject("presentation", { accentColor: COLOR }),
  },
  {
    app: "presentation",
    operation: "deleteSlides",
    schema: strictAppObject("presentation", {
      slides: { type: "array", minItems: 1, maxItems: 500, items: { type: "integer", minimum: 1 } },
      from: { type: "integer", minimum: 1 },
      to: { type: "integer", minimum: 1 },
      start: { type: "integer", minimum: 1 },
      end: { type: "integer", minimum: 1 },
    }),
  },
  {
    app: "presentation",
    operation: "normalizeLayouts",
    schema: strictAppObject("presentation", {}),
  },
  {
    app: "presentation",
    operation: "insertChart",
    schema: strictAppObject("presentation", {
      chartType: CHART_TYPE,
      ...POSITION_PROPERTIES,
    }),
  },
  {
    app: "presentation",
    operation: "insertTable",
    schema: strictAppObject("presentation", {
      rows: { type: "integer", minimum: 1 },
      columns: { type: "integer", minimum: 1 },
      values: {
        type: "array",
        maxItems: 500,
        items: { type: "array", maxItems: 100, items: {} },
      },
      ...POSITION_PROPERTIES,
    }),
  },
  {
    app: "presentation",
    operation: "replacePictureSlot",
    required: true,
    schema: strictAppObject(
      "presentation",
      {
        imagePath: NON_EMPTY_STRING,
        shapeName: { type: "string" },
        preserveAspectRatio: { type: "boolean" },
        ...POSITION_PROPERTIES,
      },
      ["imagePath"],
    ),
  },
  {
    app: "presentation",
    operation: "alignShapes",
    schema: strictAppObject("presentation", {
      minLeft: { type: "number" },
      minTop: { type: "number" },
    }),
  },
];
