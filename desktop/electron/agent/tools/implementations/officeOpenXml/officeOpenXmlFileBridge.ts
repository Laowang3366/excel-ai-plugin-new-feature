/**
 * OfficeOpenXmlFileBridge — 文件级 Office 编辑桥
 *
 * 将 Open XML 引擎适配到工具契约，支持 docx/pptx/xlsx 文件级编辑，不依赖 COM/PowerShell。
 *
 * 关联模块：
 * - officeOpenXmlEngine.ts: 实际 ZIP/XML 读写。
 * - tools/contracts/office.ts: OfficeFileBridge 契约。
 */

import type { OfficeFileBridge } from "../../contracts/office";
import { inspectOfficeOpenXmlFile, replaceOfficeOpenXmlText } from "./officeOpenXmlEngine";
import { inspectOfficeOpenXmlLayout } from "./layoutInspector";
import { inspectOfficeOpenXmlTables } from "./tableInspector";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";
import { createOfficeVisualSnapshot } from "./visualSnapshot";
import type {
  OfficeVisualSnapshotInput,
  OfficeOpenXmlLayoutInspectInput,
  OfficeOpenXmlReplaceInput,
  OfficeOpenXmlTableInspectInput,
  OfficeOpenXmlTableStyleInput,
} from "./types";

export class OfficeOpenXmlFileBridge implements OfficeFileBridge {
  inspectFile(filePath: string): Promise<unknown> {
    return inspectOfficeOpenXmlFile(filePath);
  }

  replaceText(input: OfficeOpenXmlReplaceInput): Promise<unknown> {
    return replaceOfficeOpenXmlText(input);
  }

  inspectLayout(input: OfficeOpenXmlLayoutInspectInput): Promise<unknown> {
    return inspectOfficeOpenXmlLayout(input);
  }

  inspectTable(input: OfficeOpenXmlTableInspectInput): Promise<unknown> {
    return inspectOfficeOpenXmlTables(input);
  }

  applyTableStyle(input: OfficeOpenXmlTableStyleInput): Promise<unknown> {
    return applyOfficeOpenXmlTableStyle(input);
  }

  snapshot(input: OfficeVisualSnapshotInput): Promise<unknown> {
    return createOfficeVisualSnapshot(input);
  }
}
