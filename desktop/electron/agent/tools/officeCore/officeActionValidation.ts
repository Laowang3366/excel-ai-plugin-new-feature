import { existsSync } from "node:fs";

import type { OfficeActionInput, OfficeActionResult } from "./types";

export function withValidationChecks(
  input: OfficeActionInput,
  result: OfficeActionResult,
): OfficeActionResult {
  if (result.status !== "done") return result;
  const checks: Array<{ name: string; ok: boolean; message: string }> = [];
  if (input.filePath) {
    const ok = existsSync(input.filePath);
    checks.push({
      name: "file-exists",
      ok,
      message: ok ? `文件存在: ${input.filePath}` : `文件不存在: ${input.filePath}`,
    });
  }

  const containsText = stringArrayParam(input.params?.containsText);
  if (containsText.length > 0) {
    const content = JSON.stringify(result.data ?? "").toLocaleLowerCase();
    for (const expected of containsText) {
      const ok = content.includes(expected.toLocaleLowerCase());
      checks.push({
        name: "contains-text",
        ok,
        message: ok ? `结果包含文本: ${expected}` : `结果不包含文本: ${expected}`,
      });
    }
  }

  const countPath = stringParam(input.params, "countPath");
  if (countPath) {
    const rawCount = valueAtPath(result.data, countPath);
    const count = Array.isArray(rawCount) ? rawCount.length : Number(rawCount);
    const expectedCount = numberParam(input.params, "expectedCount");
    const minCount = numberParam(input.params, "minCount");
    if (expectedCount !== undefined) {
      const ok = Number.isFinite(count) && count === expectedCount;
      checks.push({
        name: "expected-count",
        ok,
        message: ok
          ? `${countPath} 数量为 ${count}`
          : `${countPath} 数量 ${count}，预期 ${expectedCount}`,
      });
    } else if (minCount !== undefined) {
      const ok = Number.isFinite(count) && count >= minCount;
      checks.push({
        name: "minimum-count",
        ok,
        message: ok
          ? `${countPath} 数量 ${count} 不小于 ${minCount}`
          : `${countPath} 数量 ${count}，小于 ${minCount}`,
      });
    }
  }

  if (input.params?.outputExists === true) {
    const outputPath =
      result.outputPath || extractString(result.data, "outputPath") || input.outputPath;
    const ok = Boolean(outputPath && existsSync(outputPath));
    checks.push({
      name: "output-exists",
      ok,
      message: ok ? `输出文件存在: ${outputPath}` : `输出文件不存在: ${outputPath || "未返回路径"}`,
    });
  }

  if (checks.length === 0) {
    const ok = result.data !== undefined && result.data !== null;
    checks.push({
      name: "inspection-data",
      ok,
      message: ok ? "已取得检查数据" : "未取得检查数据",
    });
  }
  const ok = checks.every((check) => check.ok);
  return {
    ...result,
    summary: ok ? "Office 验证通过" : "Office 验证未通过",
    validation: { ok, checks },
  };
}

function valueAtPath(value: unknown, pathExpression: string): unknown {
  return pathExpression
    .split(".")
    .filter(Boolean)
    .reduce<unknown>((current, segment) => {
      if (Array.isArray(current) && /^\d+$/.test(segment)) return current[Number(segment)];
      return current && typeof current === "object"
        ? (current as Record<string, unknown>)[segment]
        : undefined;
    }, value);
}

function stringArrayParam(value: unknown): string[] {
  if (typeof value === "string" && value.length > 0) return [value];
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === "string" && item.length > 0)
    : [];
}

function numberParam(
  params: Record<string, unknown> | undefined,
  name: string,
): number | undefined {
  const value = params?.[name];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function extractString(value: unknown, key: string): string | undefined {
  return value &&
    typeof value === "object" &&
    key in value &&
    typeof (value as Record<string, unknown>)[key] === "string"
    ? (value as Record<string, string>)[key]
    : undefined;
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
}
