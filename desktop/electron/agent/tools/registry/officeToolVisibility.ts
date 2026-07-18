import type { ToolDefinition } from "../../shared/types";
import type { PromptRoutingContext } from "../../prompts/promptRouting";
import { resolveOfficeAdvancedIntents } from "../../prompts/promptRouting";
import {
  APPLY_OPERATIONS,
  INSPECT_OPERATIONS,
  PIVOT_OPERATIONS,
  POWER_QUERY_OPERATIONS,
  withOfficeOperationDiscriminator,
  WORKFLOW_OPERATIONS,
} from "./officeActionSchemas";

export function filterToolDefinitionsForTurn(
  definitions: ToolDefinition[],
  context: PromptRoutingContext = {},
): ToolDefinition[] {
  const intents = resolveOfficeAdvancedIntents(context);
  const allowPowerQuery = intents.has("refreshable-etl");
  const allowPivot = intents.has("interactive-pivot");

  return definitions.map((definition) => {
    if (definition.name === "office.action.apply") {
      return withOperationVisibility(
        definition,
        filterAdvancedOperations(APPLY_OPERATIONS, allowPowerQuery, allowPivot),
        buildParamsDescription(allowPowerQuery, allowPivot),
      );
    }
    if (
      definition.name === "office.action.inspect" ||
      definition.name === "office.action.validate"
    ) {
      return withOperationVisibility(
        definition,
        filterAdvancedOperations(INSPECT_OPERATIONS, allowPowerQuery, allowPivot),
      );
    }
    if (definition.name === "office.workflow.run") {
      return withWorkflowOperationVisibility(
        definition,
        filterAdvancedOperations(WORKFLOW_OPERATIONS, allowPowerQuery, allowPivot),
        buildParamsDescription(allowPowerQuery, allowPivot),
      );
    }
    if (definition.name === "office.workflow.template.save") {
      return withWorkflowOperationVisibility(
        definition,
        filterAdvancedOperations(WORKFLOW_OPERATIONS, allowPowerQuery, allowPivot),
        buildParamsDescription(allowPowerQuery, allowPivot),
      );
    }
    return definition;
  });
}

export function compactToolDefinitionsForTurn(
  definitions: ToolDefinition[],
  context: PromptRoutingContext = {},
): ToolDefinition[] {
  return filterToolDefinitionsForTurn(definitions, context).map(compactOperationDiscriminators);
}

function withOperationVisibility(
  definition: ToolDefinition,
  operations: string[],
  paramsDescription?: string,
): ToolDefinition {
  const properties = asSchema(definition.parameters.properties);
  const operation = asSchema(properties.operation);
  const params = asSchema(properties.params);
  return {
    ...definition,
    parameters: withOfficeOperationDiscriminator(
      {
        ...definition.parameters,
        properties: {
          ...properties,
          operation: {
            ...operation,
            type: "string",
            enum: operations,
            description: `本轮允许的 operation：${operations.join("、")}`,
          },
          ...(paramsDescription ? { params: { ...params, description: paramsDescription } } : {}),
        },
      },
      operations,
    ),
  };
}

function withWorkflowOperationVisibility(
  definition: ToolDefinition,
  operations: string[],
  paramsDescription: string,
): ToolDefinition {
  const properties = asSchema(definition.parameters.properties);
  const steps = asSchema(properties.steps);
  const item = asSchema(steps.items);
  const itemProperties = asSchema(item.properties);
  const operation = asSchema(itemProperties.operation);
  const params = asSchema(itemProperties.params);
  return {
    ...definition,
    parameters: {
      ...definition.parameters,
      properties: {
        ...properties,
        steps: {
          ...steps,
          items: withOfficeOperationDiscriminator(
            {
              ...item,
              properties: {
                ...itemProperties,
                operation: {
                  ...operation,
                  type: "string",
                  enum: operations,
                  description: `本轮工作流允许的 operation：${operations.join("、")}`,
                },
                params: { ...params, description: paramsDescription },
              },
            },
            operations,
          ),
        },
      },
    },
  };
}

function filterAdvancedOperations(
  operations: string[],
  allowPowerQuery: boolean,
  allowPivot: boolean,
): string[] {
  return operations.filter((operation) => {
    if (!allowPowerQuery && POWER_QUERY_OPERATIONS.has(operation)) return false;
    if (!allowPivot && PIVOT_OPERATIONS.has(operation)) return false;
    return true;
  });
}

function buildParamsDescription(allowPowerQuery: boolean, allowPivot: boolean): string {
  const boundaries = ["仅传当前 operation 明确需要的参数"];
  if (allowPowerQuery) {
    boundaries.push(
      "Power Query 需 advancedIntent:'refreshable-etl'；创建/更新另需 sourceKind:'external'|'multi-source'",
    );
  }
  if (allowPivot) {
    boundaries.push("透视表/切片器需 advancedIntent:'interactive-pivot'");
  }
  return boundaries.join("；");
}

function compactOperationDiscriminators(definition: ToolDefinition): ToolDefinition {
  if (
    definition.name === "office.action.apply" ||
    definition.name === "office.action.inspect" ||
    definition.name === "office.action.validate"
  ) {
    const parameters = { ...definition.parameters };
    const properties = asSchema(parameters.properties);
    const params = asSchema(properties.params);
    const catalog = buildCompactOperationCatalog(operationVariants(parameters));
    parameters.properties = {
      ...properties,
      params: {
        ...params,
        description: buildCompactParamsDescription(params.description, catalog, "顶层调用"),
      },
    };
    delete parameters.oneOf;
    return { ...definition, parameters };
  }
  if (
    definition.name === "office.workflow.run" ||
    definition.name === "office.workflow.template.save"
  ) {
    const properties = asSchema(definition.parameters.properties);
    const steps = asSchema(properties.steps);
    const items = { ...asSchema(steps.items) };
    const itemProperties = asSchema(items.properties);
    const params = asSchema(itemProperties.params);
    const catalog = buildCompactOperationCatalog(operationVariants(items));
    items.properties = {
      ...itemProperties,
      params: {
        ...params,
        description: buildCompactParamsDescription(params.description, catalog, "每个 steps[]"),
      },
    };
    delete items.oneOf;
    return {
      ...definition,
      parameters: {
        ...definition.parameters,
        properties: {
          ...properties,
          steps: { ...steps, items },
        },
      },
    };
  }
  return definition;
}

function buildCompactParamsDescription(
  baseDescription: unknown,
  catalog: string[],
  scope: string,
): string {
  const parts = [
    `${scope}中 operation 专属字段必须放在 params 对象内；target 是顶层定位字段，不能代替 slides、from、to 等 operation 参数`,
    "字段名后 ? 表示可选，未标 ? 表示必填；params? 表示整个 params 可省略",
  ];
  if (typeof baseDescription === "string" && baseDescription.trim()) {
    parts.push(baseDescription.trim());
  }
  if (catalog.length > 0) {
    parts.push(`operation 参数格式：\n${catalog.join("\n")}`);
  }
  return parts.join("；");
}

function buildCompactOperationCatalog(variants: Record<string, any>[]): string[] {
  const entries = new Set<string>();
  for (const variant of variants) {
    const properties = asSchema(variant.properties);
    const operations = schemaValues(properties.operation);
    const apps = schemaValues(properties.app);
    const required = new Set(Array.isArray(variant.required) ? variant.required : []);
    const paramsPrefix = required.has("params") ? "params" : "params?";
    const signature = summarizeSchema(properties.params);
    for (const operation of operations) {
      const labels = apps.length > 0 ? apps.map((app) => `${app}.${operation}`) : [operation];
      for (const label of labels) entries.add(`${label}: ${paramsPrefix} ${signature}`);
    }
  }
  return Array.from(entries).sort((left, right) => left.localeCompare(right, "en"));
}

function operationVariants(schema: Record<string, any>): Record<string, any>[] {
  return Array.isArray(schema.oneOf)
    ? schema.oneOf.filter((variant: unknown): variant is Record<string, any> =>
        Boolean(variant && typeof variant === "object" && !Array.isArray(variant)),
      )
    : [];
}

function schemaValues(schemaValue: unknown): string[] {
  const schema = asSchema(schemaValue);
  if (typeof schema.const === "string") return [schema.const];
  return Array.isArray(schema.enum)
    ? schema.enum.filter((value: unknown): value is string => typeof value === "string")
    : [];
}

function summarizeSchema(schemaValue: unknown, depth = 0): string {
  const schema = asSchema(schemaValue);
  if (Object.prototype.hasOwnProperty.call(schema, "const")) return JSON.stringify(schema.const);
  if (Array.isArray(schema.enum) && schema.enum.length > 0) {
    return schema.enum.map((value: unknown) => JSON.stringify(value)).join("|");
  }

  for (const keyword of ["oneOf", "anyOf"] as const) {
    if (Array.isArray(schema[keyword]) && schema[keyword].length > 0) {
      return uniqueSummaries(schema[keyword], depth)
        .map((summary) => (summary.includes("|") ? `(${summary})` : summary))
        .join("|");
    }
  }
  if (Array.isArray(schema.allOf) && schema.allOf.length > 0) {
    return uniqueSummaries(schema.allOf, depth).join(" & ");
  }

  if (schema.type === "array") {
    const item = summarizeSchema(schema.items, depth + 1);
    const bounds = summarizeArrayBounds(schema);
    return `${item.startsWith("{") ? item : `(${item})`}[]${bounds}`;
  }

  const properties = asSchema(schema.properties);
  if (schema.type === "object" || Object.keys(properties).length > 0) {
    if (depth >= 8) return "{...}";
    const required = new Set(Array.isArray(schema.required) ? schema.required : []);
    const fields = Object.entries(properties).map(
      ([name, propertySchema]) =>
        `${name}${required.has(name) ? "" : "?"}:${summarizeSchema(propertySchema, depth + 1)}`,
    );
    if (fields.length === 0) return schema.additionalProperties === false ? "{}" : "object";
    return `{${fields.join(",")}}`;
  }

  if (Array.isArray(schema.type)) return schema.type.join("|");
  if (typeof schema.type === "string") return summarizeScalar(schema);
  return "any";
}

function uniqueSummaries(schemas: unknown[], depth: number): string[] {
  return Array.from(new Set(schemas.map((schema) => summarizeSchema(schema, depth + 1))));
}

function summarizeScalar(schema: Record<string, any>): string {
  const type = schema.type as string;
  if (type === "string" && schema.minLength === 1) return "non-empty string";
  const bounds: string[] = [];
  if (typeof schema.minimum === "number") bounds.push(`>=${schema.minimum}`);
  if (typeof schema.maximum === "number") bounds.push(`<=${schema.maximum}`);
  return `${type}${bounds.length > 0 ? bounds.join("") : ""}`;
}

function summarizeArrayBounds(schema: Record<string, any>): string {
  const bounds: string[] = [];
  if (typeof schema.minItems === "number") bounds.push(`>=${schema.minItems}`);
  if (typeof schema.maxItems === "number") bounds.push(`<=${schema.maxItems}`);
  return bounds.length > 0 ? `[${bounds.join(",")}]` : "";
}

function asSchema(value: unknown): Record<string, any> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, any>)
    : {};
}
