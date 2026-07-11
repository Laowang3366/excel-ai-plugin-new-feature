import { doneResult } from "../../officeCore/results";
import type {
  OfficeActionApp,
  OfficeActionKind,
  OfficeActionResult,
  OfficeActionValidation,
} from "../../officeCore/types";

interface OpenXmlActionInput {
  operation: string;
  filePath: string;
  outputPath?: string;
  target?: string;
  action?: OfficeActionKind;
}

export function createOpenXmlDoneResult(app: OfficeActionApp) {
  return (
    input: OpenXmlActionInput,
    outputPath: string,
    changedParts: string[],
    summary: string,
    data?: unknown,
    validation?: OfficeActionValidation,
  ): OfficeActionResult => doneResult({
    engine: "openxml",
    app,
    action: input.action || "edit",
    operation: input.operation,
    filePath: input.filePath,
    outputPath,
    target: input.target,
    summary,
    data,
    validation: validation || {
      ok: true,
      checks: [{ name: "output-file", ok: true, message: "已生成输出文件" }],
    },
    changes: changedParts.map((partName) => ({
      kind: "openxml-part",
      target: partName,
      detail: `已更新 ${partName}`,
    })),
  });
}
