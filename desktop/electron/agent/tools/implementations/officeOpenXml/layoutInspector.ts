/**
 * Open XML 布局检查器
 *
 * 关联模块：
 * - types.ts: 定义布局检查输入输出。
 * - officeOpenXmlFileBridge.ts: 将本模块暴露给 OfficeFileBridge。
 */

import { readFile } from "fs/promises";
import JSZip from "jszip";
import { extractOpenXmlTextValues } from "../../../shared/openXmlText";
import {
  detectOfficeOpenXmlDocumentType,
  getOfficeOpenXmlTextTagName,
  isOfficeOpenXmlTextPart,
} from "./documentParts";
import type {
  OfficeOpenXmlLayoutInspectInput,
  OfficeOpenXmlLayoutInspectResult,
  OfficeOpenXmlLayoutObject,
} from "./types";

export async function inspectOfficeOpenXmlLayout(
  input: OfficeOpenXmlLayoutInspectInput
): Promise<OfficeOpenXmlLayoutInspectResult> {
  const documentType = detectOfficeOpenXmlDocumentType(input.filePath);
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const objects: OfficeOpenXmlLayoutObject[] = [];

  for (const partName of Object.keys(zip.files).filter((name) => isOfficeOpenXmlTextPart(documentType, name)).sort()) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    for (const text of extractOpenXmlTextValues(xml, { tagName: getOfficeOpenXmlTextTagName(documentType) }).filter(Boolean)) {
      objects.push({ type: "text", partName, text, textLength: text.length });
    }
  }

  return {
    engine: "openxml",
    operation: "inspectLayout",
    documentType,
    filePath: input.filePath,
    target: input.target,
    objectCount: objects.length,
    objects,
  };
}
