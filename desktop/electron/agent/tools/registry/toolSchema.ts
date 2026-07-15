import type { ToolDefinition } from "../../shared/types";

type JsonSchema = Record<string, unknown>;

export const TOOL_ARGUMENT_LIMITS = Object.freeze({
  maxJsonBytes: 4 * 1024 * 1024,
  maxDepth: 16,
  maxNodes: 100_000,
  maxStringChars: 1_000_000,
  maxArrayItems: 20_000,
  maxObjectKeys: 512,
});

export interface ToolArgumentValidationResult {
  args?: Record<string, unknown>;
  error?: string;
}

export function normalizeToolDefinition(definition: ToolDefinition): ToolDefinition {
  return {
    ...definition,
    parameters: normalizeSchema(definition.parameters),
  };
}

export function parseAndValidateToolArguments(
  argsJson: string,
  parameters?: Record<string, unknown>,
): ToolArgumentValidationResult {
  const source = argsJson?.trim() || "{}";
  if (Buffer.byteLength(source, "utf8") > TOOL_ARGUMENT_LIMITS.maxJsonBytes) {
    return { error: `参数 JSON 超过 ${TOOL_ARGUMENT_LIMITS.maxJsonBytes} 字节限制` };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(source);
  } catch (error) {
    return {
      error: `参数不是有效 JSON: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
  if (!isRecord(parsed)) return { error: "参数根节点必须为对象" };

  const budgetError = validateValueBudget(parsed, "$", 0, { nodes: 0 });
  if (budgetError) return { error: budgetError };
  if (parameters) {
    const schemaError = validateAgainstSchema(parsed, parameters, "$", 0);
    if (schemaError) return { error: schemaError };
  }
  return { args: parsed };
}

function normalizeSchema(schema: JsonSchema): JsonSchema {
  const normalized: JsonSchema = { ...schema };
  const properties = asSchemaMap(schema.properties);
  if (properties) {
    normalized.properties = Object.fromEntries(
      Object.entries(properties).map(([key, child]) => [key, normalizeSchema(child)]),
    );
    if (schema.additionalProperties === undefined) normalized.additionalProperties = false;
  } else if (schema.type === "object" && schema.additionalProperties === undefined) {
    normalized.additionalProperties = true;
  }
  const items = asSchema(schema.items);
  if (items) normalized.items = normalizeSchema(items);
  for (const keyword of ["allOf", "oneOf"] as const) {
    const schemas = asSchemaArray(schema[keyword]);
    if (schemas) normalized[keyword] = schemas.map(normalizeSchema);
  }
  return normalized;
}

function validateAgainstSchema(
  value: unknown,
  schema: JsonSchema,
  path: string,
  depth: number,
): string | undefined {
  if (depth > TOOL_ARGUMENT_LIMITS.maxDepth) return `${path} 超过最大嵌套深度`;
  if (Object.prototype.hasOwnProperty.call(schema, "const") && !Object.is(schema.const, value)) {
    return `${path} 必须为 ${String(schema.const)}`;
  }
  const enumValues = Array.isArray(schema.enum) ? schema.enum : undefined;
  if (enumValues && !enumValues.some((candidate) => Object.is(candidate, value))) {
    return `${path} 必须为 ${enumValues.map(String).join("、")} 之一`;
  }

  let typeError: string | undefined;
  switch (schema.type) {
    case undefined:
      break;
    case "string":
      if (typeof value !== "string") return `${path} 应为字符串`;
      typeError = validateStringBounds(value, schema, path);
      break;
    case "number":
      if (
        typeof value !== "number" ||
        !Number.isFinite(value) ||
        Math.abs(value) > Number.MAX_SAFE_INTEGER
      ) {
        return `${path} 应为有限安全范围内的数字`;
      }
      typeError = validateNumberBounds(value, schema, path);
      break;
    case "integer":
      if (typeof value !== "number" || !Number.isSafeInteger(value)) return `${path} 应为安全整数`;
      typeError = validateNumberBounds(value, schema, path);
      break;
    case "boolean":
      if (typeof value !== "boolean") return `${path} 应为布尔值`;
      break;
    case "array":
      if (!Array.isArray(value)) return `${path} 应为数组`;
      typeError = validateArray(value, schema, path, depth);
      break;
    case "object":
      if (!isRecord(value)) return `${path} 应为对象`;
      typeError = validateObject(value, schema, path, depth);
      break;
    default:
      return `${path} 使用了不支持的 Schema 类型: ${String(schema.type)}`;
  }
  if (typeError) return typeError;
  const allOf = asSchemaArray(schema.allOf);
  if (allOf) {
    for (const child of allOf) {
      const error = validateAgainstSchema(value, child, path, depth + 1);
      if (error) return error;
    }
  }
  const oneOf = asSchemaArray(schema.oneOf);
  if (oneOf) {
    const errors = oneOf.map((child) => validateAgainstSchema(value, child, path, depth + 1));
    const matches = errors.filter((error) => error === undefined).length;
    if (matches === 0) {
      const detail = errors.find((error) => error && !isDiscriminatorMismatch(error, path)) ?? errors[0];
      return `${path} 不匹配任何允许的参数结构: ${detail}`;
    }
    if (matches > 1) return `${path} 同时匹配多个互斥参数结构`;
  }
  return undefined;
}

function isDiscriminatorMismatch(error: string, path: string): boolean {
  return error.startsWith(`${path}.operation 必须为 `) || error.startsWith(`${path}.app 必须为 `);
}

function validateObject(
  value: Record<string, unknown>,
  schema: JsonSchema,
  path: string,
  depth: number,
): string | undefined {
  const properties = asSchemaMap(schema.properties) ?? {};
  const required = Array.isArray(schema.required)
    ? schema.required.filter((item): item is string => typeof item === "string")
    : [];
  for (const key of required) {
    if (!Object.prototype.hasOwnProperty.call(value, key) || value[key] === undefined) {
      return `${path}.${key} 为必填参数`;
    }
  }
  for (const [key, childValue] of Object.entries(value)) {
    const childSchema = properties[key];
    if (!childSchema) {
      if (schema.additionalProperties === false) return `${path}.${key} 是未声明参数`;
      continue;
    }
    const error = validateAgainstSchema(childValue, childSchema, `${path}.${key}`, depth + 1);
    if (error) return error;
  }
  return undefined;
}

function validateArray(
  value: unknown[],
  schema: JsonSchema,
  path: string,
  depth: number,
): string | undefined {
  const minItems = finiteNumber(schema.minItems);
  const maxItems = finiteNumber(schema.maxItems);
  if (minItems !== undefined && value.length < minItems) return `${path} 至少需要 ${minItems} 项`;
  if (maxItems !== undefined && value.length > maxItems) return `${path} 最多允许 ${maxItems} 项`;
  const itemSchema = asSchema(schema.items);
  if (!itemSchema) return undefined;
  for (let index = 0; index < value.length; index += 1) {
    const error = validateAgainstSchema(value[index], itemSchema, `${path}[${index}]`, depth + 1);
    if (error) return error;
  }
  return undefined;
}

function validateStringBounds(value: string, schema: JsonSchema, path: string): string | undefined {
  const minLength = finiteNumber(schema.minLength);
  const maxLength = finiteNumber(schema.maxLength);
  if (minLength !== undefined && value.length < minLength)
    return `${path} 长度不能小于 ${minLength}`;
  if (maxLength !== undefined && value.length > maxLength)
    return `${path} 长度不能超过 ${maxLength}`;
  return undefined;
}

function validateNumberBounds(value: number, schema: JsonSchema, path: string): string | undefined {
  const minimum = finiteNumber(schema.minimum);
  const maximum = finiteNumber(schema.maximum);
  if (minimum !== undefined && value < minimum) return `${path} 不能小于 ${minimum}`;
  if (maximum !== undefined && value > maximum) return `${path} 不能大于 ${maximum}`;
  return undefined;
}

function validateValueBudget(
  value: unknown,
  path: string,
  depth: number,
  state: { nodes: number },
): string | undefined {
  if (depth > TOOL_ARGUMENT_LIMITS.maxDepth)
    return `${path} 超过最大嵌套深度 ${TOOL_ARGUMENT_LIMITS.maxDepth}`;
  state.nodes += 1;
  if (state.nodes > TOOL_ARGUMENT_LIMITS.maxNodes)
    return `参数节点数超过 ${TOOL_ARGUMENT_LIMITS.maxNodes}`;
  if (typeof value === "string" && value.length > TOOL_ARGUMENT_LIMITS.maxStringChars) {
    return `${path} 字符数超过 ${TOOL_ARGUMENT_LIMITS.maxStringChars}`;
  }
  if (
    typeof value === "number" &&
    (!Number.isFinite(value) || Math.abs(value) > Number.MAX_SAFE_INTEGER)
  ) {
    return `${path} 应为有限安全范围内的数字`;
  }
  if (Array.isArray(value)) {
    if (value.length > TOOL_ARGUMENT_LIMITS.maxArrayItems) {
      return `${path} 数组项数超过 ${TOOL_ARGUMENT_LIMITS.maxArrayItems}`;
    }
    for (let index = 0; index < value.length; index += 1) {
      const error = validateValueBudget(value[index], `${path}[${index}]`, depth + 1, state);
      if (error) return error;
    }
  } else if (isRecord(value)) {
    const entries = Object.entries(value);
    if (entries.length > TOOL_ARGUMENT_LIMITS.maxObjectKeys) {
      return `${path} 对象字段数超过 ${TOOL_ARGUMENT_LIMITS.maxObjectKeys}`;
    }
    for (const [key, child] of entries) {
      if (key === "__proto__" || key === "prototype" || key === "constructor") {
        return `${path}.${key} 是禁止的保留字段`;
      }
      const error = validateValueBudget(child, `${path}.${key}`, depth + 1, state);
      if (error) return error;
    }
  }
  return undefined;
}

function asSchema(value: unknown): JsonSchema | undefined {
  return isRecord(value) ? value : undefined;
}

function asSchemaMap(value: unknown): Record<string, JsonSchema> | undefined {
  if (!isRecord(value)) return undefined;
  const entries = Object.entries(value);
  if (!entries.every(([, child]) => isRecord(child))) return undefined;
  return Object.fromEntries(entries) as Record<string, JsonSchema>;
}

function asSchemaArray(value: unknown): JsonSchema[] | undefined {
  return Array.isArray(value) && value.every(isRecord) ? value as JsonSchema[] : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function finiteNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}
