import type { JsonSchema, OfficeOperationParamSchema } from "./officeOperationParamSchemas";

const NON_EMPTY_STRING: JsonSchema = { type: "string", minLength: 1 };
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
const VALIDATION_PROPERTIES: Record<string, JsonSchema> = {
  containsText: {
    oneOf: [{ type: "string" }, { type: "array", maxItems: 256, items: { type: "string" } }],
  },
  countPath: NON_EMPTY_STRING,
  expectedCount: { type: "integer", minimum: 0 },
  minCount: { type: "integer", minimum: 0 },
  outputExists: { type: "boolean" },
};
const ANIMATION_RULE_PROPERTIES: Record<string, JsonSchema> = {
  trigger: { type: "string", enum: ["onClick", "withPrevious", "afterPrevious"] },
  order: { type: "integer", minimum: 1, maximum: 10_000 },
  duration: { type: "number", minimum: 0, maximum: 86_400 },
  delay: { type: "number", minimum: 0, maximum: 86_400 },
  repeatCount: { type: "number", minimum: 0, maximum: 10_000 },
  pathX: { type: "number", minimum: -10, maximum: 10 },
  pathY: { type: "number", minimum: -10, maximum: 10 },
};
const ANIMATION_RULE: JsonSchema = {
  oneOf: [
    ...["entrance", "exit", "path"].flatMap((category) =>
      animationRuleSchemas(category, ["fade", "appear", "fly", "dissolve", "wipe", "zoom"]),
    ),
    ...animationRuleSchemas("emphasis", ["pulse", "spin", "transparency"]),
  ],
};
const NOTE: JsonSchema = strictObject(
  {
    slideIndex: { type: "integer", minimum: 1 },
    text: { type: "string", maxLength: 1_000_000 },
    append: { type: "boolean" },
  },
  ["slideIndex", "text"],
);

export const PRESENTATION_PLAYBACK_PARAM_SCHEMAS: OfficeOperationParamSchema[] = [
  inspectionSchema("inspectPresentationTheme", false),
  inspectionSchema("inspectSlideElements", true),
  inspectionSchema("inspectAnimations", true),
  inspectionSchema("inspectSpeakerNotes", true),
  {
    app: "presentation",
    operation: "configureAnimations",
    required: true,
    schema: strictObject(
      {
        ...PRESENTATION_BASE_PROPERTIES,
        clearExisting: { type: "boolean" },
        effects: {
          type: "array",
          minItems: 1,
          maxItems: 256,
          items: ANIMATION_RULE,
        },
      },
      ["effects"],
    ),
  },
  {
    app: "presentation",
    operation: "configureSlideShow",
    required: true,
    schema: strictObject(
      {
        ...PRESENTATION_BASE_PROPERTIES,
        showType: { type: "string", enum: ["speaker", "window", "kiosk"] },
        autoPlay: { type: "boolean" },
        useSlideTimings: { type: "boolean" },
        loop: { type: "boolean" },
        showWithAnimation: { type: "boolean" },
        allSlides: { type: "boolean" },
        transition: { type: "string", enum: ["fade", "cut", "dissolve", "wipe", "none"] },
        advanceOnClick: { type: "boolean" },
        advanceAfter: { type: "number", minimum: 0, maximum: 86_400 },
        transitionDuration: { type: "number", minimum: 0, maximum: 3_600 },
      },
      ["showType"],
    ),
  },
  speakerNotesSchema(
    { text: { type: "string", maxLength: 1_000_000 }, append: { type: "boolean" } },
    ["text"],
  ),
  speakerNotesSchema(
    {
      notesBySlide: { type: "array", minItems: 1, maxItems: 500, items: NOTE },
    },
    ["notesBySlide"],
  ),
  {
    app: "presentation",
    operation: "exportHandouts",
    required: true,
    schema: strictObject(
      {
        ...PRESENTATION_BASE_PROPERTIES,
        includeNotes: { type: "boolean" },
        layout: {
          type: "string",
          enum: ["notes", "one", "two", "three", "four", "six", "nine", "outline"],
        },
      },
      ["layout"],
    ),
  },
];

function inspectionSchema(
  operation: string,
  supportsSlideSelection: boolean,
): OfficeOperationParamSchema {
  return {
    app: "presentation",
    operation,
    schema: strictObject({
      ...PRESENTATION_BASE_PROPERTIES,
      ...VALIDATION_PROPERTIES,
      ...(supportsSlideSelection ? { allSlides: { type: "boolean" } } : {}),
    }),
  };
}

function speakerNotesSchema(
  properties: Record<string, JsonSchema>,
  required: string[],
): OfficeOperationParamSchema {
  return {
    app: "presentation",
    operation: "setSpeakerNotes",
    required: true,
    schema: strictObject({ ...PRESENTATION_BASE_PROPERTIES, ...properties }, required),
  };
}

function animationRuleSchemas(category: string, effects: string[]): JsonSchema[] {
  const properties = {
    ...ANIMATION_RULE_PROPERTIES,
    category: { type: "string", const: category },
    effect: { type: "string", enum: effects },
  };
  return [
    strictObject({ ...properties, shapeName: NON_EMPTY_STRING }, [
      "category",
      "effect",
      "shapeName",
    ]),
    strictObject(
      {
        ...properties,
        shapeNames: {
          type: "array",
          minItems: 1,
          maxItems: 256,
          items: NON_EMPTY_STRING,
        },
      },
      ["category", "effect", "shapeNames"],
    ),
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
