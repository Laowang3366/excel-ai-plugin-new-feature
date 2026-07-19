import type {
  HostAdapter,
  PageOrder,
  PageOrientation,
  PagePaperSize,
  SheetPageLayoutUpdateInput,
} from "../host/types";
import type { ToolCall, ToolResult } from "./types";

const PAPER_SIZES: PagePaperSize[] = ["a3", "a4", "a5", "letter", "legal"];
const PAGE_ORDERS: PageOrder[] = ["downThenOver", "overThenDown"];

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (typeof args[key] !== "string" || (args[key] as string).trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return (args[key] as string).trim();
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) return undefined;
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalOrientation(args: Record<string, unknown>): PageOrientation | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "orientation") || args.orientation === undefined) {
    return undefined;
  }
  const value = args.orientation;
  if (value !== "portrait" && value !== "landscape") {
    throw new Error("orientation must be portrait|landscape");
  }
  return value;
}

function optionalPageOrder(args: Record<string, unknown>): PageOrder | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "pageOrder") || args.pageOrder === undefined) {
    return undefined;
  }
  const value = args.pageOrder;
  if (typeof value !== "string" || !(PAGE_ORDERS as string[]).includes(value)) {
    throw new Error("pageOrder must be downThenOver|overThenDown");
  }
  return value as PageOrder;
}

function optionalFirstPageNumber(args: Record<string, unknown>): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "firstPageNumber")) return undefined;
  const value = args.firstPageNumber;
  if (
    value === undefined ||
    value === null ||
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 1
  ) {
    throw new Error("firstPageNumber must be a finite integer >= 1");
  }
  return value;
}

function optionalZoomScale(args: Record<string, unknown>): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "zoomScale") || args.zoomScale === undefined) {
    return undefined;
  }
  const value = args.zoomScale;
  if (typeof value !== "number" || !Number.isFinite(value) || value < 10 || value > 400) {
    throw new Error("zoomScale must be a finite number between 10 and 400");
  }
  return value;
}

function optionalPaperSize(args: Record<string, unknown>): PagePaperSize | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "paperSize")) return undefined;
  const value = args.paperSize;
  if (value === undefined || value === null) {
    throw new Error("paperSize must be a3|a4|a5|letter|legal");
  }
  if (typeof value !== "string" || !(PAPER_SIZES as string[]).includes(value)) {
    throw new Error("paperSize must be a3|a4|a5|letter|legal");
  }
  return value as PagePaperSize;
}

function optionalFitPages(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  const value = args[key];
  if (
    value === undefined ||
    value === null ||
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 1 ||
    value > 32767
  ) {
    throw new Error(`${key} must be a finite integer between 1 and 32767`);
  }
  return value;
}

/** Non-empty string only; null/empty rejected (clear not a proven Office.js contract). */
function optionalNonEmptyString(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) return undefined;
  if (typeof args[key] !== "string" || (args[key] as string).trim() === "") {
    throw new Error(`${key} must be a non-empty string (clear unsupported)`);
  }
  return (args[key] as string).trim();
}

function optionalMargins(
  args: Record<string, unknown>,
): SheetPageLayoutUpdateInput["margins"] | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "margins") || args.margins === undefined) {
    return undefined;
  }
  if (args.margins === null || typeof args.margins !== "object" || Array.isArray(args.margins)) {
    throw new Error("margins must be an object");
  }
  const raw = args.margins as Record<string, unknown>;
  for (const key of Object.keys(raw)) {
    if (!["top", "bottom", "left", "right", "header", "footer"].includes(key)) {
      throw new Error(`unknown margins field: ${key}`);
    }
  }
  const out: NonNullable<SheetPageLayoutUpdateInput["margins"]> = {};
  for (const key of ["top", "bottom", "left", "right", "header", "footer"] as const) {
    if (!Object.prototype.hasOwnProperty.call(raw, key) || raw[key] === undefined) continue;
    if (typeof raw[key] !== "number" || !Number.isFinite(raw[key] as number) || (raw[key] as number) < 0) {
      throw new Error(`margins.${key} must be a non-negative finite number`);
    }
    out[key] = raw[key] as number;
  }
  if (Object.keys(out).length === 0) throw new Error("margins requires at least one side");
  return out;
}


function optionalTextSides(
  args: Record<string, unknown>,
  key: "headers" | "footers",
): Partial<{ left: string; center: string; right: string }> | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  const raw = args[key];
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error(`${key} must be an object`);
  }
  const obj = raw as Record<string, unknown>;
  for (const side of Object.keys(obj)) {
    if (!["left", "center", "right"].includes(side)) {
      throw new Error(`unknown ${key} field: ${side}`);
    }
  }
  const out: Partial<{ left: string; center: string; right: string }> = {};
  for (const side of ["left", "center", "right"] as const) {
    if (!Object.prototype.hasOwnProperty.call(obj, side) || obj[side] === undefined) continue;
    if (typeof obj[side] !== "string") {
      throw new Error(`${key}.${side} must be a string`);
    }
    out[side] = obj[side] as string;
  }
  if (Object.keys(out).length === 0) {
    throw new Error(`${key} requires at least one of left|center|right`);
  }
  return out;
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

export async function executePageLayoutTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "sheet.pageLayout.get") {
    rejectUnknown(call.arguments, ["sheetName"]);
    return fromHost(
      call.name,
      await host.getSheetPageLayout(requireString(call.arguments, "sheetName")),
    );
  }
  if (call.name === "sheet.pageLayout.set") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "orientation",
      "centerHorizontally",
      "centerVertically",
      "printGridlines",
      "printHeadings",
      "blackAndWhite",
      "draft",
      "pageOrder",
      "firstPageNumber",
      "margins",
      "headers",
      "footers",
      "zoomScale",
      "paperSize",
      "fitToPagesWide",
      "fitToPagesTall",
      "printArea",
      "printTitleRows",
      "printTitleColumns",
    ]);
    const input: SheetPageLayoutUpdateInput = {
      sheetName: requireString(call.arguments, "sheetName"),
      orientation: optionalOrientation(call.arguments),
      centerHorizontally: optionalBoolean(call.arguments, "centerHorizontally"),
      centerVertically: optionalBoolean(call.arguments, "centerVertically"),
      printGridlines: optionalBoolean(call.arguments, "printGridlines"),
      printHeadings: optionalBoolean(call.arguments, "printHeadings"),
      blackAndWhite: optionalBoolean(call.arguments, "blackAndWhite"),
      draft: optionalBoolean(call.arguments, "draft"),
      pageOrder: optionalPageOrder(call.arguments),
      firstPageNumber: optionalFirstPageNumber(call.arguments),
      margins: optionalMargins(call.arguments),
      headers: optionalTextSides(call.arguments, "headers"),
      footers: optionalTextSides(call.arguments, "footers"),
      zoomScale: optionalZoomScale(call.arguments),
      paperSize: optionalPaperSize(call.arguments),
      fitToPagesWide: optionalFitPages(call.arguments, "fitToPagesWide"),
      fitToPagesTall: optionalFitPages(call.arguments, "fitToPagesTall"),
      printArea: optionalNonEmptyString(call.arguments, "printArea"),
      printTitleRows: optionalNonEmptyString(call.arguments, "printTitleRows"),
      printTitleColumns: optionalNonEmptyString(call.arguments, "printTitleColumns"),
    };
    if (
      input.zoomScale !== undefined &&
      (input.fitToPagesWide !== undefined || input.fitToPagesTall !== undefined)
    ) {
      throw new Error("zoomScale is mutually exclusive with fitToPagesWide/fitToPagesTall");
    }
    if (
      input.orientation === undefined &&
      input.centerHorizontally === undefined &&
      input.centerVertically === undefined &&
      input.printGridlines === undefined &&
      input.printHeadings === undefined &&
      input.blackAndWhite === undefined &&
      input.draft === undefined &&
      input.pageOrder === undefined &&
      input.firstPageNumber === undefined &&
      input.margins === undefined &&
      input.headers === undefined &&
      input.footers === undefined &&
      input.zoomScale === undefined &&
      input.paperSize === undefined &&
      input.fitToPagesWide === undefined &&
      input.fitToPagesTall === undefined &&
      input.printArea === undefined &&
      input.printTitleRows === undefined &&
      input.printTitleColumns === undefined
    ) {
      throw new Error("sheet.pageLayout.set requires at least one update field");
    }
    return fromHost(call.name, await host.setSheetPageLayout(input));
  }
  return null;
}
