/**
 * Open XML Office 文件编辑类型
 *
 * 关联模块：
 * - officeOpenXmlEngine.ts: 使用这些结构读写 docx/pptx 文件。
 * - tools/executors/officeExecutors.ts: 将结构化结果返回给模型和前端预览栏。
 */

export type OfficeOpenXmlDocumentType = "word" | "presentation" | "spreadsheet";

export interface OfficeOpenXmlTextPart {
  partName: string;
  text: string;
  textLength: number;
}

export interface OfficeOpenXmlInspectResult {
  engine: "openxml";
  operation: "inspect";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  textPartCount: number;
  textCharCount: number;
  textPreview: string;
  textParts: OfficeOpenXmlTextPart[];
}

export interface OfficeOpenXmlReplaceInput {
  filePath: string;
  findText: string;
  replaceText: string;
  outputPath?: string;
  matchCase?: boolean;
}

export interface OfficeOpenXmlChangedPart {
  partName: string;
  replacements: number;
}

export interface OfficeOpenXmlReplaceResult {
  engine: "openxml";
  operation: "replaceText";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  outputPath: string;
  findText: string;
  replaceText: string;
  replacements: number;
  changedParts: OfficeOpenXmlChangedPart[];
}

export interface OfficeOpenXmlLayoutInspectInput {
  filePath: string;
  target?: string;
}

export interface OfficeOpenXmlLayoutObject {
  type: "text";
  partName: string;
  text: string;
  textLength: number;
}

export interface OfficeOpenXmlLayoutInspectResult {
  engine: "openxml";
  operation: "inspectLayout";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  target?: string;
  objectCount: number;
  objects: OfficeOpenXmlLayoutObject[];
}

export interface OfficeOpenXmlTableInspectInput {
  filePath: string;
  target?: string;
}

export interface OfficeOpenXmlTableCell {
  text: string;
  rowIndex: number;
  columnIndex: number;
  reference?: string;
  bold?: boolean;
  fillColor?: string;
  alignment?: string;
}

export interface OfficeOpenXmlTableRow {
  rowIndex: number;
  isHeaderGuess: boolean;
  cells: OfficeOpenXmlTableCell[];
}

export interface OfficeOpenXmlTableSummary {
  index: number;
  partName: string;
  rows: OfficeOpenXmlTableRow[];
  columns: number;
}

export interface OfficeOpenXmlTableInspectResult {
  engine: "openxml";
  operation: "inspectTable";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  target?: string;
  tableCount: number;
  tables: OfficeOpenXmlTableSummary[];
}

export type OfficeOpenXmlTableStylePreset = "professional" | "compact" | "financial";

export interface OfficeOpenXmlTableStyleInput {
  filePath: string;
  style: OfficeOpenXmlTableStylePreset;
  outputPath?: string;
  target?: string;
}

export interface OfficeOpenXmlTableStyleResult {
  engine: "openxml";
  operation: "applyTableStyle";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  outputPath: string;
  target?: string;
  style: OfficeOpenXmlTableStylePreset;
  changedParts: string[];
}

export interface OfficeVisualSnapshotInput {
  filePath: string;
  target?: string;
  outputPath?: string;
  preferEngine?: "openxml" | "com";
}

export interface OfficeVisualSnapshotResult {
  engine: "openxml" | "com";
  operation: "snapshot";
  documentType: OfficeOpenXmlDocumentType;
  filePath: string;
  outputPath: string;
  target?: string;
  renderer?: string;
  supported: boolean;
  error?: string;
}
