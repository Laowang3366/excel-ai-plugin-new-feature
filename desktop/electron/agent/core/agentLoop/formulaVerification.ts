import type { ToolCallItem, ToolResultItem, Turn } from "../../shared/types";
import {
  findReadEvidence,
  getLatestFormulaPreparation,
  matrixFromReadResult,
  parseFormulaRangeRef,
  parseFormulaTaskContract,
  sameFormulaAnchor,
  type FormulaAcceptanceCheck,
  type FormulaRangeRef,
} from "./formulaTaskContract";

export type FormulaVerificationStatus = "passed" | "passed_with_assumptions" | "failed";

export interface FormulaVerificationCheckResult {
  type: string;
  status: "passed" | "failed" | "skipped";
  required: boolean;
  message: string;
  evidence?: unknown;
}

export interface FormulaVerificationReport {
  status: FormulaVerificationStatus;
  writeToolCallId: string;
  referenceMode: "none" | "partial" | "complete";
  anchor: string;
  actualRange: string;
  actualShape: { rows: number; columns: number };
  checks: FormulaVerificationCheckResult[];
  errorCells: Array<{ row: number; column: number; value: string }>;
  sampleMismatches: Array<{ row: number; column: number; expected: unknown; actual: unknown }>;
  assumptions: string[];
  summary: string;
  nextActions: string[];
}

interface FormulaWriteEvidence {
  call: ToolCallItem;
  result: ToolResultItem;
  ref: FormulaRangeRef;
  callIndex: number;
  resultIndex: number;
}

interface FormulaValidationReadEvidence {
  call: ToolCallItem;
  result: ToolResultItem;
  ref: FormulaRangeRef;
  matrix: unknown[][];
  resultIndex: number;
}

export function buildPendingFormulaValidationRead(turn: Turn): Record<string, unknown> | null {
  const write = getLatestSuccessfulFormulaWrite(turn);
  if (!write || hasVerificationForWrite(turn, write.call.id)) return null;
  if (findValidationRead(turn, write)) return null;
  return {
    sheetName: write.ref.sheetName,
    range: write.ref.range,
    expand: "spill",
  };
}

export function shouldRunFormulaVerification(turn: Turn): boolean {
  const write = getLatestSuccessfulFormulaWrite(turn);
  return Boolean(write && !hasVerificationForWrite(turn, write.call.id) && findValidationRead(turn, write));
}

export function verifyLatestFormulaWrite(turn: Turn): FormulaVerificationReport | string {
  const contract = parseFormulaTaskContract(turn);
  if (!contract) return "当前任务缺少公式任务契约";
  const preparation = getLatestFormulaPreparation(turn)?.preparation;
  if (!preparation || preparation.status !== "ready") return "公式解题与验收契约尚未准备完成";
  const write = getLatestSuccessfulFormulaWrite(turn);
  if (!write) return "尚未找到成功的公式写入";
  const validationRead = findValidationRead(turn, write);
  if (!validationRead) return "尚未从公式锚点回读实际结果";

  const output = validationRead.matrix;
  const checks: FormulaVerificationCheckResult[] = [];
  const errorCells = collectExcelErrors(output);
  checks.push({
    type: "no_excel_error",
    status: errorCells.length === 0 ? "passed" : "failed",
    required: true,
    message: errorCells.length === 0
      ? "实际结果未发现 Excel/WPS 错误值"
      : `实际结果包含 ${errorCells.length} 个 Excel/WPS 错误值`,
    evidence: errorCells.slice(0, 20),
  });

  const sampleMismatches: FormulaVerificationReport["sampleMismatches"] = [];
  if (contract.referenceRange) {
    const reference = findReadEvidence(turn, contract.referenceRange)?.matrix;
    if (!reference) {
      checks.push({
        type: "reference_sample",
        status: "failed",
        required: true,
        message: "指定的参考样例没有可用读取结果",
      });
    } else {
      const comparison = compareSample(output, reference, contract.referenceMode === "complete");
      sampleMismatches.push(...comparison.mismatches);
      checks.push({
        type: "reference_sample",
        status: comparison.passed ? "passed" : "failed",
        required: true,
        message: comparison.message,
        evidence: comparison.mismatches.slice(0, 20),
      });
    }
  }

  for (const check of preparation.acceptanceChecks) {
    if (check.type === "no_excel_error") continue;
    checks.push(runAcceptanceCheck(check, output, turn));
  }

  const failedRequired = checks.filter((check) => check.required && check.status !== "passed");
  const status: FormulaVerificationStatus = failedRequired.length > 0
    ? "failed"
    : preparation.assumptions.length > 0
      ? "passed_with_assumptions"
      : "passed";
  const nextActions = failedRequired.map((check) => check.message).slice(0, 6);
  return {
    status,
    writeToolCallId: write.call.id,
    referenceMode: contract.referenceMode,
    anchor: write.ref.raw,
    actualRange: validationRead.ref.raw,
    actualShape: matrixShape(output),
    checks,
    errorCells: errorCells.slice(0, 20),
    sampleMismatches: sampleMismatches.slice(0, 20),
    assumptions: preparation.assumptions,
    summary: status === "failed"
      ? `公式已写入，但 ${failedRequired.length} 项必需验收未通过。`
      : status === "passed_with_assumptions"
        ? "公式实际结果通过验收，并记录了次要假设。"
        : "公式实际结果通过验收。",
    nextActions,
  };
}

export function getLatestFormulaVerification(
  turn: Turn,
  writeToolCallId?: string,
): FormulaVerificationReport | null {
  for (let index = turn.items.length - 1; index >= 0; index--) {
    const item = turn.items[index];
    if (item.type !== "tool_result" || item.toolName !== "formula.verify" || item.isError) continue;
    if (!isVerificationReport(item.result)) continue;
    if (writeToolCallId && item.result.writeToolCallId !== writeToolCallId) continue;
    return item.result;
  }
  return null;
}

export function getLatestSuccessfulFormulaWrite(turn: Turn): FormulaWriteEvidence | null {
  const calls = new Map<string, { call: ToolCallItem; index: number }>();
  turn.items.forEach((item, index) => {
    if (item.type === "tool_call" && item.toolName === "range.write") {
      calls.set(item.id, { call: item, index });
    }
  });
  for (let resultIndex = turn.items.length - 1; resultIndex >= 0; resultIndex--) {
    const item = turn.items[resultIndex];
    if (item.type !== "tool_result" || item.toolName !== "range.write" || item.isError) continue;
    const match = calls.get(item.toolCallId);
    if (!match) continue;
    const sheetName = text(match.call.arguments.sheetName) || undefined;
    const range = text(match.call.arguments.range);
    if (!range) continue;
    return {
      call: match.call,
      result: item,
      ref: { raw: sheetName ? `${sheetName}!${range}` : range, sheetName, range },
      callIndex: match.index,
      resultIndex,
    };
  }
  return null;
}

function findValidationRead(
  turn: Turn,
  write: FormulaWriteEvidence,
): FormulaValidationReadEvidence | null {
  const calls = new Map<string, ToolCallItem>();
  turn.items.forEach((item) => {
    if (item.type === "tool_call" && item.toolName === "range.read") calls.set(item.id, item);
  });
  for (let resultIndex = turn.items.length - 1; resultIndex > write.resultIndex; resultIndex--) {
    const item = turn.items[resultIndex];
    if (item.type !== "tool_result" || item.toolName !== "range.read" || item.isError) continue;
    const call = calls.get(item.toolCallId);
    if (!call || call.arguments.expand !== "spill") continue;
    const sheetName = text(call.arguments.sheetName) || undefined;
    const requestedRange = text(call.arguments.range);
    const actualAddress = isRecord(item.result) ? text(item.result.address) : "";
    const ref = parseFormulaRangeRef(`${sheetName ? `${sheetName}!` : ""}${actualAddress || requestedRange}`);
    if (!ref || !sameFormulaAnchor(ref, write.ref)) continue;
    return {
      call,
      result: item,
      ref,
      matrix: matrixFromReadResult(item.result) ?? [],
      resultIndex,
    };
  }
  return null;
}

function hasVerificationForWrite(turn: Turn, writeToolCallId: string): boolean {
  return turn.items.some((item) =>
    item.type === "tool_result"
    && item.toolName === "formula.verify"
    && !item.isError
    && isRecord(item.result)
    && item.result.writeToolCallId === writeToolCallId,
  );
}

function runAcceptanceCheck(
  check: FormulaAcceptanceCheck,
  output: unknown[][],
  turn: Turn,
): FormulaVerificationCheckResult {
  try {
    switch (check.type) {
      case "shape":
        return checkShape(check, output);
      case "unique_key":
        return checkUniqueKey(check, output);
      case "row_count":
        return checkRowCount(check, output, turn);
      case "aggregate_reconciliation":
        return checkAggregate(check, output, turn);
      case "sort_order":
        return checkSortOrder(check, output);
      case "lookup_consistency":
        return checkLookupConsistency(check, output, turn, false);
      case "pattern_match":
        return checkPattern(check, output);
      case "boundary":
        return checkBoundary(check, output);
      case "spot_check":
        return checkLookupConsistency(check, output, turn, true);
      default:
        return skipped(check, `暂不支持自动执行 ${check.type} 检查`);
    }
  } catch (error) {
    return {
      type: check.type,
      status: check.required ? "failed" : "skipped",
      required: check.required,
      message: `${check.description}：检查参数无效或执行失败`,
      evidence: error instanceof Error ? error.message : String(error),
    };
  }
}

function checkShape(check: FormulaAcceptanceCheck, output: unknown[][]): FormulaVerificationCheckResult {
  const shape = matrixShape(output);
  const expectedRows = positiveInt(check.params.expectedRows);
  const expectedColumns = positiveInt(check.params.expectedColumns);
  const minRows = positiveInt(check.params.minRows);
  const minColumns = positiveInt(check.params.minColumns);
  const passed = (expectedRows === null || shape.rows === expectedRows)
    && (expectedColumns === null || shape.columns === expectedColumns)
    && (minRows === null || shape.rows >= minRows)
    && (minColumns === null || shape.columns >= minColumns);
  return result(check, passed, passed ? check.description : `${check.description}：实际为 ${shape.rows} 行 ${shape.columns} 列`, shape);
}

function checkUniqueKey(check: FormulaAcceptanceCheck, output: unknown[][]): FormulaVerificationCheckResult {
  const headerRows = nonNegativeInt(check.params.headerRows) ?? 1;
  const columns = oneBasedColumns(check.params.outputColumns, [1]);
  const keys = output.slice(headerRows).map((row) => columns.map((column) => normalizeCell(row[column - 1])).join("\u001f"));
  const nonBlank = keys.filter((key) => key.replace(/\u001f/g, "") !== "");
  const duplicateCount = nonBlank.length - new Set(nonBlank).size;
  return result(
    check,
    duplicateCount === 0,
    duplicateCount === 0 ? check.description : `${check.description}：发现 ${duplicateCount} 个重复业务键`,
    { checkedRows: nonBlank.length, duplicateCount },
  );
}

function checkRowCount(check: FormulaAcceptanceCheck, output: unknown[][], turn: Turn): FormulaVerificationCheckResult {
  const outputHeaderRows = nonNegativeInt(check.params.outputHeaderRows) ?? 1;
  const actual = Math.max(0, output.length - outputHeaderRows);
  const mode = text(check.params.mode) || "exact";
  let expected: number | null = nonNegativeInt(check.params.expectedRows);
  if (mode === "unique_source") {
    const source = sourceMatrix(turn, check.params.sourceRange);
    if (!source) throw new Error("找不到 row_count 指定的数据源读取结果");
    const sourceHeaderRows = nonNegativeInt(check.params.sourceHeaderRows) ?? 1;
    const columns = oneBasedColumns(check.params.sourceColumns, [1]);
    const keys = source.slice(sourceHeaderRows)
      .map((row) => columns.map((column) => normalizeCell(row[column - 1])).join("\u001f"))
      .filter((key) => key.replace(/\u001f/g, "") !== "");
    expected = new Set(keys).size;
  } else if (mode === "sum_source") {
    const source = sourceMatrix(turn, check.params.sourceRange);
    if (!source) throw new Error("找不到 row_count 指定的数据源读取结果");
    const sourceHeaderRows = nonNegativeInt(check.params.sourceHeaderRows) ?? 1;
    const sourceColumn = positiveInt(check.params.sourceColumn) ?? 1;
    expected = Math.round(sumColumn(source, sourceColumn, sourceHeaderRows));
  }
  if (expected === null) throw new Error("row_count 缺少 expectedRows 或可计算的数据源规则");
  return result(check, actual === expected, actual === expected ? check.description : `${check.description}：预期 ${expected} 行，实际 ${actual} 行`, { expected, actual });
}

function checkAggregate(check: FormulaAcceptanceCheck, output: unknown[][], turn: Turn): FormulaVerificationCheckResult {
  const source = sourceMatrix(turn, check.params.sourceRange);
  if (!source) throw new Error("找不到 aggregate_reconciliation 指定的数据源读取结果");
  const sourceColumn = positiveInt(check.params.sourceColumn);
  const outputColumn = positiveInt(check.params.outputColumn);
  if (!sourceColumn || !outputColumn) throw new Error("缺少 sourceColumn 或 outputColumn");
  const sourceTotal = sumColumn(source, sourceColumn, nonNegativeInt(check.params.sourceHeaderRows) ?? 1);
  const outputTotal = sumColumn(output, outputColumn, nonNegativeInt(check.params.outputHeaderRows) ?? 1);
  const tolerance = nonNegativeNumber(check.params.tolerance) ?? 0.000001;
  const passed = Math.abs(sourceTotal - outputTotal) <= tolerance;
  return result(check, passed, passed ? check.description : `${check.description}：源数据合计 ${sourceTotal}，输出合计 ${outputTotal}`, { sourceTotal, outputTotal, tolerance });
}

function checkSortOrder(check: FormulaAcceptanceCheck, output: unknown[][]): FormulaVerificationCheckResult {
  const column = positiveInt(check.params.outputColumn) ?? 1;
  const direction = text(check.params.direction).toLowerCase() === "asc" ? "asc" : "desc";
  const values = output.slice(nonNegativeInt(check.params.headerRows) ?? 1)
    .map((row) => row[column - 1])
    .filter((value) => normalizeCell(value) !== "");
  let passed = true;
  for (let index = 1; index < values.length; index++) {
    const comparison = compareCells(values[index - 1], values[index]);
    if ((direction === "asc" && comparison > 0) || (direction === "desc" && comparison < 0)) {
      passed = false;
      break;
    }
  }
  return result(check, passed, passed ? check.description : `${check.description}：第 ${column} 列未按 ${direction} 排列`, { direction, checkedRows: values.length });
}

function checkLookupConsistency(
  check: FormulaAcceptanceCheck,
  output: unknown[][],
  turn: Turn,
  sampleOnly: boolean,
): FormulaVerificationCheckResult {
  const source = sourceMatrix(turn, check.params.sourceRange);
  if (!source) throw new Error("找不到 lookup/spot_check 指定的数据源读取结果");
  const sourceKeyColumn = positiveInt(check.params.sourceKeyColumn) ?? 1;
  const sourceValueColumn = positiveInt(check.params.sourceValueColumn) ?? 2;
  const outputKeyColumn = positiveInt(check.params.outputKeyColumn) ?? 1;
  const outputValueColumn = positiveInt(check.params.outputValueColumn) ?? 2;
  const sourceHeaderRows = nonNegativeInt(check.params.sourceHeaderRows) ?? 1;
  const outputHeaderRows = nonNegativeInt(check.params.outputHeaderRows) ?? 1;
  const lookup = new Map<string, unknown>();
  for (const row of source.slice(sourceHeaderRows)) {
    const key = normalizeCell(row[sourceKeyColumn - 1]);
    if (key && !lookup.has(key)) lookup.set(key, row[sourceValueColumn - 1]);
  }
  const rows = output.slice(outputHeaderRows);
  const selected = sampleOnly ? pickRepresentativeRows(rows, positiveInt(check.params.sampleSize) ?? 5) : rows;
  const mismatches: Array<{ key: string; expected: unknown; actual: unknown }> = [];
  for (const row of selected) {
    const key = normalizeCell(row[outputKeyColumn - 1]);
    if (!key) continue;
    if (!lookup.has(key)) {
      mismatches.push({ key, expected: "数据源中存在该业务键", actual: "数据源中未找到" });
      continue;
    }
    const expected = lookup.get(key);
    const actual = row[outputValueColumn - 1];
    if (!cellsEqual(expected, actual)) mismatches.push({ key, expected, actual });
  }
  return result(check, mismatches.length === 0, mismatches.length === 0 ? check.description : `${check.description}：发现 ${mismatches.length} 个键值不一致`, mismatches.slice(0, 20));
}

function checkPattern(check: FormulaAcceptanceCheck, output: unknown[][]): FormulaVerificationCheckResult {
  const column = positiveInt(check.params.outputColumn) ?? 1;
  const pattern = text(check.params.pattern);
  if (!pattern || pattern.length > 500) throw new Error("pattern 为空或过长");
  const regex = new RegExp(pattern, text(check.params.flags).replace(/[^gimsuy]/g, ""));
  const values = output.slice(nonNegativeInt(check.params.headerRows) ?? 1, (nonNegativeInt(check.params.headerRows) ?? 1) + 1_000)
    .map((row) => normalizeCell(row[column - 1]).slice(0, 500))
    .filter(Boolean);
  const invalid = values.filter((value) => {
    regex.lastIndex = 0;
    return !regex.test(value);
  });
  return result(check, invalid.length === 0, invalid.length === 0 ? check.description : `${check.description}：${invalid.length} 个值不符合格式`, invalid.slice(0, 20));
}

function checkBoundary(check: FormulaAcceptanceCheck, output: unknown[][]): FormulaVerificationCheckResult {
  const minRows = nonNegativeInt(check.params.minRows) ?? 1;
  const allowBlank = check.params.allowBlank !== false;
  const data = output.slice(nonNegativeInt(check.params.headerRows) ?? 0);
  const blankCount = data.flat().filter((value) => normalizeCell(value) === "").length;
  const passed = data.length >= minRows && (allowBlank || blankCount === 0);
  return result(check, passed, passed ? check.description : `${check.description}：边界要求未满足`, { rows: data.length, blankCount, allowBlank });
}

function sourceMatrix(turn: Turn, value: unknown): unknown[][] | null {
  const raw = text(value);
  if (!raw) return null;
  const ref = parseFormulaRangeRef(raw);
  return ref ? findReadEvidence(turn, ref)?.matrix ?? null : null;
}

function compareSample(output: unknown[][], reference: unknown[][], exactShape: boolean): {
  passed: boolean;
  message: string;
  mismatches: FormulaVerificationReport["sampleMismatches"];
} {
  const outputShape = matrixShape(output);
  const referenceShape = matrixShape(reference);
  const mismatches: FormulaVerificationReport["sampleMismatches"] = [];
  if (exactShape && (outputShape.rows !== referenceShape.rows || outputShape.columns !== referenceShape.columns)) {
    return {
      passed: false,
      message: `完整样例尺寸为 ${referenceShape.rows}x${referenceShape.columns}，实际为 ${outputShape.rows}x${outputShape.columns}`,
      mismatches,
    };
  }
  if (outputShape.rows < referenceShape.rows || outputShape.columns < referenceShape.columns) {
    return {
      passed: false,
      message: `实际结果 ${outputShape.rows}x${outputShape.columns} 小于参考样例 ${referenceShape.rows}x${referenceShape.columns}`,
      mismatches,
    };
  }
  for (let row = 0; row < referenceShape.rows; row++) {
    for (let column = 0; column < referenceShape.columns; column++) {
      if (!cellsEqual(reference[row]?.[column], output[row]?.[column])) {
        mismatches.push({ row: row + 1, column: column + 1, expected: reference[row]?.[column], actual: output[row]?.[column] });
      }
    }
  }
  return {
    passed: mismatches.length === 0,
    message: mismatches.length === 0
      ? exactShape ? "实际结果与完整参考样例逐项一致" : "实际结果与部分参考样例对应区域一致"
      : `实际结果与参考样例存在 ${mismatches.length} 个差异`,
    mismatches,
  };
}

function collectExcelErrors(matrix: unknown[][]): FormulaVerificationReport["errorCells"] {
  const errors: FormulaVerificationReport["errorCells"] = [];
  const pattern = /^#(?:REF!|VALUE!|N\/A|SPILL!|NAME\?|DIV\/0!|NUM!|NULL!|CALC!|FIELD!|GETTING_DATA)$/i;
  matrix.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    const textValue = normalizeCell(value);
    if (pattern.test(textValue)) errors.push({ row: rowIndex + 1, column: columnIndex + 1, value: textValue });
  }));
  return errors;
}

function result(check: FormulaAcceptanceCheck, passed: boolean, message: string, evidence?: unknown): FormulaVerificationCheckResult {
  return { type: check.type, status: passed ? "passed" : "failed", required: check.required, message, evidence };
}

function skipped(check: FormulaAcceptanceCheck, message: string): FormulaVerificationCheckResult {
  return { type: check.type, status: "skipped", required: check.required, message };
}

function matrixShape(matrix: unknown[][]): { rows: number; columns: number } {
  return { rows: matrix.length, columns: matrix.reduce((max, row) => Math.max(max, row.length), 0) };
}

function sumColumn(matrix: unknown[][], oneBasedColumn: number, headerRows: number): number {
  return matrix.slice(headerRows).reduce((sum, row) => sum + numericValue(row[oneBasedColumn - 1]), 0);
}

function numericValue(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function cellsEqual(left: unknown, right: unknown): boolean {
  const leftNumber = parseNumericCell(left);
  const rightNumber = parseNumericCell(right);
  if (leftNumber !== null && rightNumber !== null) {
    return Math.abs(leftNumber - rightNumber) <= 0.000001;
  }
  return normalizeCell(left) === normalizeCell(right);
}

function compareCells(left: unknown, right: unknown): number {
  const leftNumber = Number(left);
  const rightNumber = Number(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) return leftNumber - rightNumber;
  return normalizeCell(left).localeCompare(normalizeCell(right), "zh-CN");
}

function normalizeCell(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function pickRepresentativeRows(rows: unknown[][], size: number): unknown[][] {
  if (size <= 1) return rows.length > 0 ? [rows[0]] : [];
  if (rows.length <= size) return rows;
  const indexes = new Set<number>([0, rows.length - 1]);
  while (indexes.size < size) {
    indexes.add(Math.round((indexes.size - 1) * (rows.length - 1) / Math.max(1, size - 1)));
  }
  return [...indexes].sort((a, b) => a - b).map((index) => rows[index]);
}

function parseNumericCell(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || value.trim() === "") return null;
  const parsed = Number(value.replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function oneBasedColumns(value: unknown, fallback: number[]): number[] {
  if (!Array.isArray(value)) return fallback;
  const columns = value.map(positiveInt).filter((item): item is number => item !== null);
  return columns.length > 0 ? columns : fallback;
}

function positiveInt(value: unknown): number | null {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : null;
}

function nonNegativeInt(value: unknown): number | null {
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : null;
}

function nonNegativeNumber(value: unknown): number | null {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function text(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isVerificationReport(value: unknown): value is FormulaVerificationReport {
  return isRecord(value)
    && (value.status === "passed" || value.status === "passed_with_assumptions" || value.status === "failed")
    && typeof value.writeToolCallId === "string";
}
