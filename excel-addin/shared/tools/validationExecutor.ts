/**
 * Conditional format + data validation tool dispatch (keeps ToolExecutor under line limit).
 */
import type { HostAdapter, HostResult } from "../host/types";
import type { ToolCall, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";
import { requireIdent } from "./argValidation";
import {
  optionalDvErrorAlert,
  optionalDvPrompt,
  requireCfRule,
  requireDvRule,
} from "./ruleValidation";

function fromHost<T>(name: ToolCall["name"], result: HostResult<T>): ToolResult {
  return mapHostResultToToolResult(name, result);
}

export async function executeValidationTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  switch (call.name) {
    case "conditionalFormat.list":
      return fromHost(
        call.name,
        await host.listConditionalFormats(
          requireIdent(call.arguments, "sheetName"),
          requireIdent(call.arguments, "range"),
        ),
      );
    case "conditionalFormat.add":
      return fromHost(
        call.name,
        await host.addConditionalFormat({
          sheetName: requireIdent(call.arguments, "sheetName"),
          range: requireIdent(call.arguments, "range"),
          rule: requireCfRule(call.arguments),
        }),
      );
    case "conditionalFormat.delete":
      return fromHost(
        call.name,
        await host.deleteConditionalFormat(
          requireIdent(call.arguments, "sheetName"),
          requireIdent(call.arguments, "range"),
          requireIdent(call.arguments, "id"),
        ),
      );
    case "dataValidation.read":
      return fromHost(
        call.name,
        await host.readDataValidation(
          requireIdent(call.arguments, "sheetName"),
          requireIdent(call.arguments, "range"),
        ),
      );
    case "dataValidation.write":
      return fromHost(
        call.name,
        await host.writeDataValidation({
          sheetName: requireIdent(call.arguments, "sheetName"),
          range: requireIdent(call.arguments, "range"),
          rule: requireDvRule(call.arguments),
          errorAlert: optionalDvErrorAlert(call.arguments),
          prompt: optionalDvPrompt(call.arguments),
        }),
      );
    case "dataValidation.clear":
      return fromHost(
        call.name,
        await host.clearDataValidation(
          requireIdent(call.arguments, "sheetName"),
          requireIdent(call.arguments, "range"),
        ),
      );
    default:
      return null;
  }
}
