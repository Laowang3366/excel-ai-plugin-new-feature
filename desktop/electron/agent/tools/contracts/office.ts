/**
 * Office 工具契约
 *
 * 被 Word、PowerPoint、Office 脚本工具执行器和具体 COM 桥实现共享。
 */

import type { OfficeActionInput, OfficeActionResult } from "../officeCore/types";

/**
 * Word 文档桥接接口
 */
export interface WordDocumentBridge {
  /** 打开已有 Word 文档 */
  openDocument(filePath: string): Promise<{ success: boolean; documentName?: string; error?: string }>;
  /** 创建新 Word 文档 */
  createDocument(filePath: string): Promise<{ success: boolean; documentName?: string; error?: string }>;
  /** 检查当前活动文档结构 */
  inspectDocument(): Promise<unknown>;
  /** 读取当前活动文档文本 */
  readText(maxChars?: number): Promise<unknown>;
  /** 插入文本 */
  insertText(text: string, position?: string): Promise<unknown>;
  /** 插入带 Word 标题样式的段落 */
  insertHeading(text: string, level?: number, position?: string): Promise<unknown>;
  /** 查找替换文本 */
  replaceText(findText: string, replaceText: string, matchCase?: boolean): Promise<unknown>;
  /** 保存当前活动文档 */
  saveDocument(saveAsPath?: string): Promise<{ success: boolean; error?: string }>;
}

/**
 * PowerPoint 演示文稿桥接接口
 */
export interface PresentationBridge {
  /** 打开已有演示文稿 */
  openPresentation(filePath: string): Promise<{ success: boolean; presentationName?: string; error?: string }>;
  /** 创建新演示文稿 */
  createPresentation(filePath: string): Promise<{ success: boolean; presentationName?: string; error?: string }>;
  /** 检查当前活动演示文稿结构 */
  inspectPresentation(): Promise<unknown>;
  /** 读取指定幻灯片文本 */
  readSlide(slideIndex: number): Promise<unknown>;
  /** 添加幻灯片 */
  addSlide(title?: string, body?: string, layout?: string): Promise<unknown>;
  /** 设置形状文本 */
  setShapeText(slideIndex: number, text: string, shapeName?: string, shapeIndex?: number): Promise<unknown>;
  /** 在全部幻灯片文本形状中查找替换文本 */
  replaceText(findText: string, replaceText: string, matchCase?: boolean): Promise<unknown>;
  /** 保存当前活动演示文稿 */
  savePresentation(saveAsPath?: string): Promise<{ success: boolean; error?: string }>;
}

/**
 * Office Open XML 文件桥接接口
 */
export interface OfficeFileBridge {
  /** 检查 docx/pptx/xlsx 文件结构和文本摘要 */
  inspectFile(filePath: string): Promise<unknown>;
  /** 在 docx/pptx/xlsx 文件内查找替换文本 */
  replaceText(input: {
    filePath: string;
    findText: string;
    replaceText: string;
    outputPath?: string;
    matchCase?: boolean;
  }): Promise<unknown>;
  /** 检查 docx/pptx/xlsx 文件中的布局对象和基础样式 */
  inspectLayout(input: { filePath: string; target?: string }): Promise<unknown>;
  /** 检查 docx/pptx/xlsx 文件中的表格结构和样式信号 */
  inspectTable(input: { filePath: string; target?: string }): Promise<unknown>;
  /** 使用 Open XML 向表格应用保守样式预设 */
  applyTableStyle(input: {
    filePath: string;
    style: "professional" | "compact" | "financial";
    outputPath?: string;
    target?: string;
  }): Promise<unknown>;
  /** 为页面、幻灯片、工作表或表格区域创建视觉快照 */
  snapshot(input: {
    filePath: string;
    target?: string;
    outputPath?: string;
    preferEngine?: "openxml" | "com";
  }): Promise<unknown>;
}

/**
 * Office 通用脚本桥接接口
 */
export interface OfficeScriptBridge {
  /** 在指定 Office 应用 COM 对象上执行 PowerShell 脚本 */
  executeScript(app: "word" | "presentation", code: string): Promise<unknown>;
}

/**
 * 统一 Office action 桥接接口
 */
export interface OfficeActionBridge {
  /** 执行跨 Excel/Word/PPT 的统一高级操作 */
  executeAction(input: OfficeActionInput): Promise<OfficeActionResult>;
}
