import type { HostAdapter } from "../host/types";
import { isGeometricShapeType } from "../host/shapeTypes";
import type { GeometricShapeType, ShapeCreateInput } from "../host/shapeTypes";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function optionalTrimmed(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = (args[key] as string).trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

function optionalText(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "text") || args.text === undefined) {
    return undefined;
  }
  if (args.text === null) throw new Error("text must not be null");
  if (typeof args.text !== "string") throw new Error("Invalid string argument: text");
  return args.text;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalFinite(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
  }
  return value;
}

function optionalPositive(args: Record<string, unknown>, key: string): number | undefined {
  const value = optionalFinite(args, key);
  if (value != null && value <= 0) throw new Error(`${key} must be > 0`);
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
  if (result.ok) return { ok: true, tool, data: result.data };
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      error: result.reason ?? "host failed",
      detail: result,
      unsupported: true,
    };
  }
  return { ok: false, tool, error: result.reason ?? "host failed", detail: result };
}

function parseCreate(args: Record<string, unknown>): ShapeCreateInput {
  rejectUnknown(args, [
    "sheetName",
    "kind",
    "geometricType",
    "text",
    "name",
    "left",
    "top",
    "width",
    "height",
  ]);
  const sheetName = requireString(args, "sheetName");
  const kind = args.kind;
  if (kind !== "geometric" && kind !== "textBox") {
    throw new Error("kind must be geometric|textBox");
  }
  const name = optionalTrimmed(args, "name");
  const left = optionalFinite(args, "left");
  const top = optionalFinite(args, "top");
  const width = optionalPositive(args, "width");
  const height = optionalPositive(args, "height");
  if (kind === "geometric") {
    if (!Object.prototype.hasOwnProperty.call(args, "geometricType") || args.geometricType == null) {
      throw new Error("geometricType is required for kind=geometric");
    }
    if (!isGeometricShapeType(args.geometricType)) {
      throw new Error(
        "geometricType must be rectangle|ellipse|triangle|diamond|rightArrow",
      );
    }
    if (Object.prototype.hasOwnProperty.call(args, "text") && args.text !== undefined) {
      throw new Error("text is only valid for kind=textBox");
    }
    return {
      sheetName,
      kind: "geometric",
      geometricType: args.geometricType as GeometricShapeType,
      name,
      left,
      top,
      width,
      height,
    };
  }
  // Explicit null/any geometricType on textBox is rejected (not omitted).
  if (Object.prototype.hasOwnProperty.call(args, "geometricType")) {
    throw new Error("geometricType is only valid for kind=geometric");
  }
  return {
    sheetName,
    kind: "textBox",
    text: optionalText(args),
    name,
    left,
    top,
    width,
    height,
  };
}

export async function executeShapeTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "shape.list") {
    rejectUnknown(call.arguments, ["sheetName"]);
    const sheetName = optionalTrimmed(call.arguments, "sheetName");
    return fromHost(call.name, await host.listShapes(sheetName));
  }
  if (call.name === "shape.create") {
    return fromHost(call.name, await host.createShape(parseCreate(call.arguments)));
  }
  if (call.name === "shape.delete") {
    rejectUnknown(call.arguments, ["sheetName", "shapeName"]);
    return fromHost(
      call.name,
      await host.deleteShape(
        requireString(call.arguments, "sheetName"),
        requireString(call.arguments, "shapeName"),
      ),
    );
  }
  if (call.name === "shape.update") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "shapeName",
      "newName",
      "left",
      "top",
      "width",
      "height",
      "text",
      "visible",
    ]);
    const sheetName = requireString(call.arguments, "sheetName");
    const shapeName = requireString(call.arguments, "shapeName");
    const newName = optionalTrimmed(call.arguments, "newName");
    const left = optionalFinite(call.arguments, "left");
    const top = optionalFinite(call.arguments, "top");
    const width = optionalPositive(call.arguments, "width");
    const height = optionalPositive(call.arguments, "height");
    const text = optionalText(call.arguments);
    const visible = optionalBoolean(call.arguments, "visible");
    if (
      newName === undefined &&
      left === undefined &&
      top === undefined &&
      width === undefined &&
      height === undefined &&
      text === undefined &&
      visible === undefined
    ) {
      throw new Error("shape.update requires at least one update field");
    }
    return fromHost(
      call.name,
      await host.updateShape({
        sheetName,
        shapeName,
        newName,
        left,
        top,
        width,
        height,
        text,
        visible,
      }),
    );
  }
  return null;
}
