import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import {
  normalizeToolDefinition,
  parseAndValidateToolArguments,
  TOOL_ARGUMENT_LIMITS,
} from "./toolSchema";

type Schema = Record<string, unknown>;

describe("tool schema runtime validation", () => {
  it("uses strict objects in the same schema exposed to models", () => {
    const definition = normalizeToolDefinition({
      name: "test.strict",
      description: "test",
      parameters: {
        type: "object",
        properties: {
          mode: { type: "string", enum: ["safe"] },
          nested: { type: "object", properties: { count: { type: "number" } } },
        },
        required: ["mode"],
      },
      riskLevel: "safe",
      requiresApproval: false,
    });

    expect(definition.parameters).toMatchObject({
      additionalProperties: false,
      properties: { nested: { additionalProperties: false } },
    });
    expect(
      parseAndValidateToolArguments('{"mode":"safe","nested":{"count":1}}', definition.parameters)
        .error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments('{"mode":"unsafe"}', definition.parameters).error,
    ).toContain("必须为 safe");
    expect(
      parseAndValidateToolArguments('{"mode":"safe","extra":1}', definition.parameters).error,
    ).toContain("未声明参数");
    expect(
      parseAndValidateToolArguments('{"mode":"safe","nested":{"extra":1}}', definition.parameters)
        .error,
    ).toContain("未声明参数");
  });

  it("rejects malformed JSON and generic resource-budget violations", () => {
    expect(parseAndValidateToolArguments("[").error).toContain("不是有效 JSON");
    expect(parseAndValidateToolArguments("[]").error).toContain("根节点必须为对象");
    const oversized = JSON.stringify({
      value: "x".repeat(TOOL_ARGUMENT_LIMITS.maxStringChars + 1),
    });
    expect(parseAndValidateToolArguments(oversized).error).toContain("字符数超过");
    expect(parseAndValidateToolArguments('{"value":1e400}').error).toContain("有限安全范围");
    expect(
      parseAndValidateToolArguments('{"params":{"__proto__":{"polluted":true}}}').error,
    ).toContain("禁止的保留字段");
  });

  it("supports const and mutually exclusive schemas at runtime", () => {
    const definition = normalizeToolDefinition({
      name: "test.discriminator",
      description: "test",
      parameters: {
        type: "object",
        properties: {
          operation: { type: "string" },
          params: { type: "object" },
        },
        required: ["operation", "params"],
        oneOf: [
          {
            type: "object",
            properties: {
              operation: { type: "string", const: "alpha" },
              params: { type: "object", properties: { count: { type: "integer" } }, required: ["count"] },
            },
            required: ["operation", "params"],
          },
          {
            type: "object",
            properties: {
              operation: { type: "string", const: "beta" },
              params: { type: "object", properties: { enabled: { type: "boolean" } }, required: ["enabled"] },
            },
            required: ["operation", "params"],
          },
        ],
      },
      riskLevel: "safe",
      requiresApproval: false,
    });

    expect(parseAndValidateToolArguments(
      '{"operation":"alpha","params":{"count":1}}',
      definition.parameters,
    ).error).toBeUndefined();
    expect(parseAndValidateToolArguments(
      '{"operation":"alpha","params":{"enabled":true}}',
      definition.parameters,
    ).error).toContain("不匹配任何允许的参数结构");
    expect(parseAndValidateToolArguments(
      '{"operation":"alpha","params":{"count":1,"extra":true}}',
      definition.parameters,
    ).error).toContain("未声明参数");
  });

  it("validates object key names, field counts, and schema-valued additional properties", () => {
    const definition = normalizeToolDefinition({
      name: "test.dictionary",
      description: "test",
      parameters: {
        type: "object",
        minProperties: 1,
        maxProperties: 2,
        propertyNames: {
          type: "string",
          pattern: "^[A-Za-z_][A-Za-z0-9_-]{0,15}$",
        },
        additionalProperties: { type: "integer", minimum: 0 },
      },
      riskLevel: "safe",
      requiresApproval: false,
    });

    expect(parseAndValidateToolArguments('{"alpha":1,"beta-2":0}', definition.parameters).error)
      .toBeUndefined();
    expect(parseAndValidateToolArguments("{}", definition.parameters).error).toContain(
      "至少需要 1 个字段",
    );
    expect(parseAndValidateToolArguments('{"alpha":1,"beta":2,"gamma":3}', definition.parameters).error)
      .toContain("最多允许 2 个字段");
    expect(parseAndValidateToolArguments('{"bad.name":1}', definition.parameters).error).toContain(
      "格式不符合要求",
    );
    expect(parseAndValidateToolArguments('{"alpha":"1"}', definition.parameters).error).toContain(
      "应为安全整数",
    );
  });

  it.each(ALL_TOOL_DEFINITIONS)(
    "accepts a generated valid sample and rejects unknown fields for $name",
    (definition) => {
      const sample = sampleForSchema(definition.parameters) as Record<string, unknown>;
      expectStrictDeclaredObjects(definition.parameters);
      expect(
        parseAndValidateToolArguments(JSON.stringify(sample), definition.parameters).error,
      ).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...sample, __unexpected: true }),
          definition.parameters,
        ).error,
      ).toContain("未声明参数");
      expect(parseAndValidateToolArguments("[]", definition.parameters).error).toContain(
        "根节点必须为对象",
      );

      const required = Array.isArray(definition.parameters.required)
        ? definition.parameters.required.filter((key): key is string => typeof key === "string")
        : [];
      if (required.length > 0) {
        const missing = { ...sample };
        delete missing[required[0]];
        expect(
          parseAndValidateToolArguments(JSON.stringify(missing), definition.parameters).error,
        ).toContain("必填参数");
      }

      const properties = isRecord(definition.parameters.properties)
        ? definition.parameters.properties
        : {};
      const typedEntry = Object.entries(properties).find(
        ([, property]) => isRecord(property) && typeof property.type === "string",
      );
      if (typedEntry) {
        const [key, property] = typedEntry as [string, Schema];
        expect(
          parseAndValidateToolArguments(
            JSON.stringify({ ...sample, [key]: wrongTypeValue(property.type) }),
            definition.parameters,
          ).error,
        ).toContain(`$.${key}`);
      }
      for (const [key, property] of Object.entries(properties)) {
        if (!isRecord(property) || !Array.isArray(property.enum)) continue;
        expect(
          parseAndValidateToolArguments(
            JSON.stringify({ ...sample, [key]: "__invalid_enum__" }),
            definition.parameters,
          ).error,
        ).toContain(`$.${key}`);
      }
    },
  );

  it("rejects unknown fields inside workflow step items", () => {
    const workflow = ALL_TOOL_DEFINITIONS.find(
      (definition) => definition.name === "office.workflow.run",
    );
    expect(workflow).toBeDefined();
    const args = {
      steps: [
        {
          app: "excel",
          action: "inspect",
          operation: "inspectFile",
          filePath: "C:/book.xlsx",
          unknownStepField: true,
        },
      ],
    };
    expect(
      parseAndValidateToolArguments(JSON.stringify(args), workflow?.parameters).error,
    ).toContain("unknownStepField");
  });

  it("enforces declared operation bounds before execution", () => {
    const web = definition("web.search");
    expect(
      parseAndValidateToolArguments('{"query":"news","maxResults":11}', web.parameters).error,
    ).toContain("不能大于 10");

    const ocr = definition("ocr.parseDocument");
    expect(parseAndValidateToolArguments('{"filePaths":[]}', ocr.parameters).error).toContain(
      "至少需要 1 项",
    );
    expect(
      parseAndValidateToolArguments(
        '{"filePaths":["C:/scan.pdf"],"maxTextChars":999}',
        ocr.parameters,
      ).error,
    ).toContain("不能小于 1000");

    const workflow = definition("office.workflow.run");
    expect(parseAndValidateToolArguments('{"steps":[]}', workflow.parameters).error).toContain(
      "至少需要 1 项",
    );
    const step = {
      app: "excel",
      action: "inspect",
      operation: "inspectFile",
      filePath: "C:/book.xlsx",
      timeoutMs: 4_999,
    };
    expect(
      parseAndValidateToolArguments(JSON.stringify({ steps: [step] }), workflow.parameters).error,
    ).toContain("不能小于 5000");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          steps: [{ ...step, timeoutMs: 5_000, retry: { maxAttempts: 1.5 } }],
        }),
        workflow.parameters,
      ).error,
    ).toContain("安全整数");
  });
});

function sampleForSchema(schema: Schema): unknown {
  if (Object.prototype.hasOwnProperty.call(schema, "const")) return schema.const;
  let result = sampleForPlainSchema(schema);
  const allOf = Array.isArray(schema.allOf) ? schema.allOf : [];
  for (const child of allOf) {
    if (isRecord(child)) result = mergeSamples(result, sampleForSchema(child));
  }
  const oneOf = Array.isArray(schema.oneOf) ? schema.oneOf : [];
  const selected = oneOf.find(isRecord);
  if (selected) result = mergeSamples(result, sampleForSchema(selected));
  return result;
}

function sampleForPlainSchema(schema: Schema): unknown {
  const enumValues = Array.isArray(schema.enum) ? schema.enum : [];
  if (enumValues.length > 0) return enumValues[0];
  switch (schema.type) {
    case "string":
      return "sample";
    case "number":
    case "integer":
      return typeof schema.minimum === "number" ? schema.minimum : 1;
    case "boolean":
      return true;
    case "array": {
      const minItems = typeof schema.minItems === "number" ? schema.minItems : 0;
      return Array.from({ length: minItems }, () =>
        sampleForSchema(isRecord(schema.items) ? schema.items : {}),
      );
    }
    case "object": {
      const properties = isRecord(schema.properties) ? schema.properties : {};
      const required = Array.isArray(schema.required) ? schema.required : [];
      return Object.fromEntries(
        required
          .filter((key): key is string => typeof key === "string")
          .map((key) => [key, sampleForSchema(isRecord(properties[key]) ? properties[key] : {})]),
      );
    }
    default:
      return null;
  }
}

function mergeSamples(base: unknown, selected: unknown): unknown {
  return isRecord(base) && isRecord(selected) ? { ...base, ...selected } : selected;
}

function isRecord(value: unknown): value is Schema {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function wrongTypeValue(type: unknown): unknown {
  switch (type) {
    case "string":
      return 1;
    case "number":
      return "1";
    case "integer":
      return 1.5;
    case "boolean":
      return "true";
    case "array":
      return {};
    case "object":
      return [];
    default:
      return undefined;
  }
}

function definition(name: string) {
  const result = ALL_TOOL_DEFINITIONS.find((candidate) => candidate.name === name);
  if (!result) throw new Error(`Missing tool definition: ${name}`);
  return result;
}

function expectStrictDeclaredObjects(schema: Schema): void {
  const properties = isRecord(schema.properties) ? schema.properties : undefined;
  if (properties) {
    expect(schema.additionalProperties).toBe(false);
    for (const property of Object.values(properties)) {
      if (isRecord(property)) expectStrictDeclaredObjects(property);
    }
  }
  if (isRecord(schema.additionalProperties)) {
    expectStrictDeclaredObjects(schema.additionalProperties);
  }
  if (isRecord(schema.propertyNames)) expectStrictDeclaredObjects(schema.propertyNames);
  if (isRecord(schema.items)) expectStrictDeclaredObjects(schema.items);
  for (const keyword of ["allOf", "oneOf"] as const) {
    const schemas = Array.isArray(schema[keyword]) ? schema[keyword] : [];
    for (const child of schemas) {
      if (isRecord(child)) expectStrictDeclaredObjects(child);
    }
  }
}
