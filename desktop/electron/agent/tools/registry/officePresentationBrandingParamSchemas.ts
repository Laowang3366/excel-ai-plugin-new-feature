import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
const COLOR: JsonSchema = { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" };
const ACTION_TIMEOUT: JsonSchema = {
  type: "integer",
  minimum: 5_000,
  maximum: 600_000,
};
const PRESENTATION_BASE_PROPERTIES: Record<string, JsonSchema> = {
  host: { type: "string", enum: ["powerpoint", "wps"] },
  instanceId: NON_EMPTY_STRING,
  actionTimeoutMs: ACTION_TIMEOUT,
};
const SHAPE_NAMES: JsonSchema = {
  type: "array",
  minItems: 1,
  maxItems: 500,
  items: NON_EMPTY_STRING,
};
const FONT_MAP: JsonSchema = {
  type: "object",
  maxProperties: 256,
  propertyNames: { type: "string", minLength: 1, maxLength: 255 },
  additionalProperties: { type: "string", minLength: 1, maxLength: 255 },
};
const THEME_COLOR: JsonSchema = strictObject(
  {
    index: { type: "integer", minimum: 1, maximum: 12 },
    value: COLOR,
  },
  ["index", "value"],
);
const LAYOUT_MAPPING: JsonSchema = {
  oneOf: [
    strictObject(
      {
        slideIndex: { type: "integer", minimum: 1 },
        layoutName: NON_EMPTY_STRING,
      },
      ["slideIndex", "layoutName"],
    ),
    strictObject(
      {
        slideName: NON_EMPTY_STRING,
        layoutName: NON_EMPTY_STRING,
      },
      ["slideName", "layoutName"],
    ),
  ],
};
const TABLE_CELL_EDIT: JsonSchema = strictObject(
  {
    row: { type: "integer", minimum: 1 },
    column: { type: "integer", minimum: 1 },
    text: { type: "string" },
    fontName: NON_EMPTY_STRING,
    fontSize: { type: "number", minimum: 1, maximum: 400 },
    fillColor: COLOR,
  },
  ["row", "column"],
);
const SHAPE_EDIT_PROPERTIES: Record<string, JsonSchema> = {
  preserveAspectRatio: { type: "boolean" },
  left: { type: "number", minimum: -100_000, maximum: 100_000 },
  top: { type: "number", minimum: -100_000, maximum: 100_000 },
  width: { type: "number", minimum: 1, maximum: 100_000 },
  height: { type: "number", minimum: 1, maximum: 100_000 },
  rotation: { type: "number", minimum: -360_000, maximum: 360_000 },
  text: { type: "string" },
  fontName: NON_EMPTY_STRING,
  fontSize: { type: "number", minimum: 1, maximum: 400 },
  tableCells: { type: "array", minItems: 1, maxItems: 10_000, items: TABLE_CELL_EDIT },
  chart: strictObject({
    chartType: { type: "integer", minimum: -10_000, maximum: 10_000 },
    title: { type: "string" },
    hasLegend: { type: "boolean" },
  }),
  crop: strictObject({
    left: { type: "number", minimum: 0, maximum: 100_000 },
    right: { type: "number", minimum: 0, maximum: 100_000 },
    top: { type: "number", minimum: 0, maximum: 100_000 },
    bottom: { type: "number", minimum: 0, maximum: 100_000 },
  }),
};
const SHAPE_EDIT: JsonSchema = {
  oneOf: [
    strictObject({ ...SHAPE_EDIT_PROPERTIES, shapeName: NON_EMPTY_STRING }, ["shapeName"]),
    strictObject({ ...SHAPE_EDIT_PROPERTIES, shapeIndex: { type: "integer", minimum: 1 } }, [
      "shapeIndex",
    ]),
  ],
};
const LAYOUT_SHARED_PROPERTIES: Record<string, JsonSchema> = {
  ...PRESENTATION_BASE_PROPERTIES,
  allSlides: { type: "boolean" },
  excludePlaceholders: { type: "boolean" },
  align: { type: "string", enum: ["left", "center", "right", "top", "middle", "bottom"] },
  distribute: { type: "string", enum: ["horizontal", "vertical"] },
  fitToSlide: { type: "boolean" },
  preserveAspectRatio: { type: "boolean" },
  edits: { type: "array", minItems: 1, maxItems: 1_000, items: SHAPE_EDIT },
};
const GRID_PROPERTIES: Record<string, JsonSchema> = {
  ...PRESENTATION_BASE_PROPERTIES,
  allSlides: { type: "boolean" },
  excludePlaceholders: { type: "boolean" },
  shapeNames: SHAPE_NAMES,
  columns: { type: "integer", minimum: 1, maximum: 100 },
  margin: { type: "number", minimum: 0, maximum: 100_000 },
  gap: { type: "number", minimum: 0, maximum: 100_000 },
  rowHeight: { type: "number", minimum: 1, maximum: 100_000 },
  resize: { type: "boolean" },
  preserveAspectRatio: { type: "boolean" },
};

export const PRESENTATION_BRANDING_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  {
    app: "presentation",
    operation: "applyMasterBranding",
    required: true,
    schema: strictObject(
      {
        ...PRESENTATION_BASE_PROPERTIES,
        templatePath: NON_EMPTY_STRING,
        backgroundColor: COLOR,
        themeColors: { type: "array", minItems: 1, maxItems: 12, items: THEME_COLOR },
        accentColor: COLOR,
        fontName: NON_EMPTY_STRING,
        fontMap: FONT_MAP,
        applyAccentToText: { type: "boolean" },
        logoPath: NON_EMPTY_STRING,
        logoWidth: { type: "number", minimum: 1, maximum: 100_000 },
        logoHeight: { type: "number", minimum: 1, maximum: 100_000 },
        logoLeft: { type: "number", minimum: -100_000, maximum: 100_000 },
        logoTop: { type: "number", minimum: -100_000, maximum: 100_000 },
        footerText: { type: "string", maxLength: 32_767 },
        showSlideNumber: { type: "boolean" },
        layoutMap: { type: "array", minItems: 1, maxItems: 500, items: LAYOUT_MAPPING },
      },
      ["showSlideNumber"],
    ),
  },
  layoutSchema(
    "precise",
    {
      ...PRESENTATION_BASE_PROPERTIES,
      allSlides: { type: "boolean" },
      edits: LAYOUT_SHARED_PROPERTIES.edits,
    },
    ["mode", "edits"],
  ),
  layoutSchema("grid", GRID_PROPERTIES, ["mode", "shapeNames"]),
  layoutSchema("auto", GRID_PROPERTIES, ["mode", "shapeNames"]),
  layoutSchema(
    "align",
    {
      ...PRESENTATION_BASE_PROPERTIES,
      allSlides: { type: "boolean" },
      excludePlaceholders: { type: "boolean" },
      shapeNames: SHAPE_NAMES,
      align: LAYOUT_SHARED_PROPERTIES.align,
    },
    ["mode", "shapeNames", "align"],
  ),
  layoutSchema(
    "distribute",
    {
      ...PRESENTATION_BASE_PROPERTIES,
      allSlides: { type: "boolean" },
      excludePlaceholders: { type: "boolean" },
      shapeNames: SHAPE_NAMES,
      distribute: LAYOUT_SHARED_PROPERTIES.distribute,
    },
    ["mode", "shapeNames", "distribute"],
  ),
  layoutSchema(
    "fit",
    {
      ...PRESENTATION_BASE_PROPERTIES,
      allSlides: { type: "boolean" },
      excludePlaceholders: { type: "boolean" },
      shapeNames: SHAPE_NAMES,
      fitToSlide: { type: "boolean", const: true },
    },
    ["mode", "shapeNames", "fitToSlide"],
  ),
];

function layoutSchema(
  mode: string,
  properties: Record<string, JsonSchema>,
  required: string[],
): OfficeOperationParamSchema {
  return {
    app: "presentation",
    operation: "layoutElements",
    required: true,
    schema: strictObject({ ...properties, mode: { type: "string", const: mode } }, required),
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
