import type { HostAdapter } from "../host/types";
import type {
  FormulaProtectionCommand,
  FormulaProtectionScope,
} from "../host/formulaProtectionTypes";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
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
  const value = args[key];
  if (typeof value !== "string") throw new Error(`${key} must be a string`);
  // password may be empty string → treat as omitted
  if (value === "") return undefined;
  return value;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`${key} must be a boolean`);
  return args[key] as boolean;
}

function requireScope(args: Record<string, unknown>): FormulaProtectionScope {
  const value = requireString(args, "scope");
  if (value !== "workbook" && value !== "sheet" && value !== "target") {
    throw new Error("scope must be workbook|sheet|target");
  }
  return value;
}

function requireCommand(args: Record<string, unknown>): FormulaProtectionCommand {
  const value = requireString(args, "command");
  if (value !== "lock" && value !== "unlock") {
    throw new Error("command must be lock|unlock");
  }
  return value;
}

function rejectUnknown(args: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function fromHost(
  tool: ToolCall["name"],
  result: { ok: boolean; data?: unknown; reason?: string; unsupported?: boolean },
): ToolResult {
  if (result.ok) {
    // Defense: never leak password-like fields if host mishandles.
    const data = stripSecrets(result.data);
    return { ok: true, tool, data };
  }
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      error: result.reason ?? "host failed",
      detail: stripSecrets(result),
      unsupported: true,
    };
  }
  return {
    ok: false,
    tool,
    error: result.reason ?? "host failed",
    detail: stripSecrets(result),
  };
}

function stripSecrets(value: unknown): unknown {
  if (value == null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map(stripSecrets);
  const out: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (/^(password|apiKey|api_key|token|secret)$/i.test(key)) continue;
    out[key] = stripSecrets(child);
  }
  return out;
}

export async function executeFormulaProtectionTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "formula.protection.inspect") {
    rejectUnknown(call.arguments, ["scope", "sheetName", "range"]);
    const scope = requireScope(call.arguments);
    const sheetName = optionalString(call.arguments, "sheetName");
    const range = optionalString(call.arguments, "range");
    if (scope !== "workbook" && !sheetName) {
      throw new Error("sheetName is required for scope sheet|target");
    }
    if (scope === "target" && !range) {
      throw new Error("range is required for scope=target");
    }
    return fromHost(
      call.name,
      await host.inspectFormulaProtection({ scope, sheetName, range }),
    );
  }
  if (call.name === "formula.protection.manage") {
    rejectUnknown(call.arguments, [
      "command",
      "scope",
      "sheetName",
      "range",
      "password",
      "unlockInputs",
      "protectSheet",
    ]);
    const command = requireCommand(call.arguments);
    const scope = requireScope(call.arguments);
    const sheetName = optionalString(call.arguments, "sheetName");
    const range = optionalString(call.arguments, "range");
    if (scope !== "workbook" && !sheetName) {
      throw new Error("sheetName is required for scope sheet|target");
    }
    if (scope === "target" && !range) {
      throw new Error("range is required for scope=target");
    }
    // password stays request-local; stripSecrets removes from any tool result.
    const password = optionalString(call.arguments, "password");
    return fromHost(
      call.name,
      await host.manageFormulaProtection({
        command,
        scope,
        sheetName,
        range,
        password,
        unlockInputs: optionalBoolean(call.arguments, "unlockInputs"),
        protectSheet: optionalBoolean(call.arguments, "protectSheet"),
      }),
    );
  }
  return null;
}
