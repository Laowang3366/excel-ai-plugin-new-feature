import type { HostAdapter } from "../host/types";
import type { FormulaWriteArgs, RangeWriteArgs, ToolResult } from "./types";

function fromHost(
  tool: "range.write" | "formula.write",
  result: { ok: boolean; data?: unknown; reason?: string; unsupported?: boolean },
): ToolResult {
  if (result.ok) {
    return { ok: true, tool, data: result.data };
  }
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      error: result.reason ?? "host failed",
      detail: result,
      unsupported: true,
    };
  }
  return {
    ok: false,
    tool,
    error: result.reason ?? "host failed",
    detail: result,
  };
}

export async function writeRangeWithVerify(
  host: HostAdapter,
  args: RangeWriteArgs,
): Promise<ToolResult> {
  const written = await host.writeRange(args.sheetName, args.range, args.values);
  if (!written.ok) return fromHost("range.write", written);
  if (!args.verify) {
    return { ok: true, tool: "range.write", data: written.data };
  }
  const verified = await host.readRange(args.sheetName, args.range);
  if (!verified.ok) {
    return {
      ok: true,
      tool: "range.write",
      data: written.data,
      verification: { ok: false, detail: verified },
    };
  }
  return {
    ok: true,
    tool: "range.write",
    data: written.data,
    verification: { ok: true, data: verified.data },
  };
}

export async function writeFormulaWithVerify(
  host: HostAdapter,
  args: FormulaWriteArgs,
): Promise<ToolResult> {
  const formula = args.formula.startsWith("=") ? args.formula : `=${args.formula}`;
  const written = await host.writeFormulas(args.sheetName, args.range, [[formula]]);
  if (!written.ok) return fromHost("formula.write", written);
  if (!args.verify) {
    return { ok: true, tool: "formula.write", data: written.data };
  }
  const verified = await host.readRange(args.sheetName, args.range);
  if (!verified.ok) {
    return {
      ok: true,
      tool: "formula.write",
      data: written.data,
      verification: { ok: false, detail: verified },
    };
  }
  return {
    ok: true,
    tool: "formula.write",
    data: {
      sheetName: verified.data.sheetName,
      address: verified.data.address,
      formulas: verified.data.formulas,
      values: verified.data.values,
    },
    verification: { ok: true, data: verified.data },
  };
}
