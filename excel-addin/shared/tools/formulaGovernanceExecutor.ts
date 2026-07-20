import type { HostAdapter } from "../host/types";
import type { FormulaGovernanceScope } from "../host/formulaGovernanceTypes";
import type { FormulaReplacement } from "../formulaGovernance";
import type { ToolCall, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (args[key] === undefined || args[key] === null) {
    throw new Error(`${key} must not be null/undefined`);
  }
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function optionalString(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "string") throw new Error(`${key} must be a string`);
  const v = (args[key] as string).trim();
  return v === "" ? undefined : v;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`${key} must be a boolean`);
  return args[key] as boolean;
}

function requireScope(args: Record<string, unknown>): FormulaGovernanceScope {
  const value = requireString(args, "scope");
  if (value !== "workbook" && value !== "sheet" && value !== "target") {
    throw new Error("scope must be workbook|sheet|target");
  }
  return value;
}

function rejectUnknown(args: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function requireReplacements(args: Record<string, unknown>): FormulaReplacement[] {
  const raw = args.replacements;
  if (!Array.isArray(raw) || raw.length === 0) {
    throw new Error("replacements must be a non-empty array");
  }
  if (raw.length > 1000) throw new Error("replacements maxItems is 1000");
  const out: FormulaReplacement[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      throw new Error("each replacement must be an object");
    }
    const rec = item as Record<string, unknown>;
    for (const k of Object.keys(rec)) {
      if (k !== "find" && k !== "replace") throw new Error(`unknown field in replacement: ${k}`);
    }
    if (typeof rec.find !== "string" || rec.find.trim() === "") {
      throw new Error("replacement.find must be a non-empty string");
    }
    if (typeof rec.replace !== "string") {
      throw new Error("replacement.replace must be a string");
    }
    out.push({ find: rec.find, replace: rec.replace });
  }
  return out;
}

/** Returns ToolResult if handled, else null. */
export async function executeFormulaGovernanceTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  const args = call.arguments ?? {};
  switch (call.name) {
    case "formula.dependencies.inspect": {
      rejectUnknown(args, ["scope", "sheetName", "range"]);
      const scope = requireScope(args);
      return mapHostResultToToolResult(
        call.name,
        await host.inspectFormulaDependencies({
          scope,
          sheetName: optionalString(args, "sheetName"),
          range: optionalString(args, "range"),
        }),
      );
    }
    case "formula.references.repair": {
      rejectUnknown(args, [
        "scope",
        "sheetName",
        "range",
        "replacements",
        "applyAllMappings",
      ]);
      const scope = requireScope(args);
      return mapHostResultToToolResult(
        call.name,
        await host.repairFormulaReferences({
          scope,
          sheetName: optionalString(args, "sheetName"),
          range: optionalString(args, "range"),
          replacements: requireReplacements(args),
          applyAllMappings: optionalBoolean(args, "applyAllMappings"),
        }),
      );
    }
    case "formula.convertToValues": {
      rejectUnknown(args, ["scope", "sheetName", "range", "createBackup", "backupId"]);
      const scope = requireScope(args);
      return mapHostResultToToolResult(
        call.name,
        await host.convertFormulasToValues({
          scope,
          sheetName: optionalString(args, "sheetName"),
          range: optionalString(args, "range"),
          createBackup: optionalBoolean(args, "createBackup"),
          backupId: optionalString(args, "backupId"),
        }),
      );
    }
    case "formula.backups.inspect": {
      rejectUnknown(args, []);
      return mapHostResultToToolResult(call.name, await host.inspectFormulaBackups());
    }
    case "formula.backups.restore": {
      rejectUnknown(args, ["backupId", "removeAfterRestore"]);
      return mapHostResultToToolResult(
        call.name,
        await host.restoreFormulas({
          backupId: requireString(args, "backupId"),
          removeAfterRestore: optionalBoolean(args, "removeAfterRestore"),
        }),
      );
    }
    default:
      return null;
  }
}
