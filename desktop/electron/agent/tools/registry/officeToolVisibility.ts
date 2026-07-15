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
    if (definition.name === "office.action.inspect" || definition.name === "office.action.validate") {
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
    return definition;
  });
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
    parameters: withOfficeOperationDiscriminator({
      ...definition.parameters,
      properties: {
        ...properties,
        operation: {
          ...operation,
          type: "string",
          enum: operations,
          description: `本轮允许的 operation：${operations.join("、")}`,
        },
        ...(paramsDescription
          ? { params: { ...params, description: paramsDescription } }
          : {}),
      },
    }, operations),
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
          items: withOfficeOperationDiscriminator({
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
          }, operations),
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

function asSchema(value: unknown): Record<string, any> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, any>
    : {};
}
