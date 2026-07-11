import path from "path";

import type { OfficeOpenXmlDocumentType } from "./types";

const TEXT_PART_PATTERNS: Record<OfficeOpenXmlDocumentType, RegExp> = {
  word: /^word\/(?:document|header\d+|footer\d+)\.xml$/,
  presentation: /^ppt\/slides\/slide\d+\.xml$/,
  spreadsheet: /^(?:xl\/sharedStrings\.xml|xl\/worksheets\/sheet\d+\.xml)$/,
};

const DOCUMENT_TYPE_BY_EXTENSION: Record<string, OfficeOpenXmlDocumentType> = {
  ".docx": "word",
  ".pptx": "presentation",
  ".xlsx": "spreadsheet",
};

export function detectOfficeOpenXmlDocumentType(filePath: string): OfficeOpenXmlDocumentType {
  const documentType = DOCUMENT_TYPE_BY_EXTENSION[path.extname(filePath).toLowerCase()];
  if (documentType) return documentType;
  throw new Error(`仅支持 .docx、.pptx 和 .xlsx 文件: ${filePath}`);
}

export function isOfficeOpenXmlTextPart(
  documentType: OfficeOpenXmlDocumentType,
  partName: string,
): boolean {
  return TEXT_PART_PATTERNS[documentType].test(partName);
}

export function getOfficeOpenXmlTextTagName(documentType: OfficeOpenXmlDocumentType): string {
  if (documentType === "word") return "w:t";
  if (documentType === "presentation") return "a:t";
  return "t";
}
