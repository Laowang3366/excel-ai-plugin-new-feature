/**
 * OCR/document visibility tool definitions.
 *
 * Converts local files into text, table rows and structure clues for models
 * without reliable multimodal input.
 */

import type { ToolDefinition } from "../../shared/types";

const OCR_PARSE_DOCUMENT_DEF: ToolDefinition = {
  name: "ocr.parseDocument",
  description: "解析本地文件的可见内容，返回 Markdown 文本、表格行和结构线索。默认先使用本地解析，仅在用户开启远程数据处理且本地无法完成时，才依次尝试已配置 token 的 MinerU 标准解析和 MinerU 免费 Agent。适合没有多模态能力的模型在抽取、总结、比对、质量判断、样式优化、修改验收等任务前先理解文件可见内容。",
  parameters: {
    type: "object",
    properties: {
      filePaths: {
        type: "array",
        description: "要识别的本地文件绝对路径列表，支持 PNG/JPG/JPEG/WebP/BMP/TIFF/PDF/DOC/DOCX/PPT/PPTX/XLS/XLSX/XLSM/CSV/MD/TXT",
        items: { type: "string" },
      },
      mode: {
        type: "string",
        enum: ["ocr", "invoice", "layout", "style"],
        description: "解析意图。ocr 为通用文字/表格识别；invoice 用于票据字段抽取前置识别；layout 用于可见结构和布局诊断；style 用于样式状态评估或修改后验收。工具返回解析文本供模型继续判断",
        default: "ocr",
      },
      maxTextChars: {
        type: "number",
        description: "返回给模型的最大解析文本字符数，默认60000，最多120000",
        default: 60000,
      },
      maxTableRows: {
        type: "number",
        description: "返回表格行的最大数量，默认200，最多1000",
        default: 200,
      },
      allowTokenMineru: {
        type: "boolean",
        description: "是否允许优先使用已配置 token 的 MinerU 标准解析，默认 true。只有用户明确要求不消耗 token 时才设为 false",
        default: true,
      },
      allowFreeMineru: {
        type: "boolean",
        description: "是否允许在标准 MinerU 不可用时使用 MinerU 免费 Agent 轻量解析，默认 true",
        default: true,
      },
      allowLocalFallback: {
        type: "boolean",
        description: "是否允许优先使用本地免费解析和内置工具兜底，默认 true",
        default: true,
      },
    },
    required: ["filePaths"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  isDataEgress: true,
};

export const OCR_TOOL_DEFINITIONS: ToolDefinition[] = [
  OCR_PARSE_DOCUMENT_DEF,
];
