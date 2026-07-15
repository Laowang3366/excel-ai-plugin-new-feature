import type { ToolDefinition } from "../../shared/types";

const WORD_OPEN_DEF: ToolDefinition = {
  name: "word.open",
  description:
    "把指定路径的 Word 文档（.doc/.docx）打开到 Word 应用窗口并设为活动文档。文件级读取/编辑优先使用 office.action.*",
  parameters: {
    type: "object",
    properties: {
      filePath: {
        type: "string",
        description: "Word 文档的绝对路径，如 C:\\Users\\用户\\Desktop\\报告.docx",
      },
    },
    required: ["filePath"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

const WORD_INSPECT_DEF: ToolDefinition = {
  name: "word.inspect",
  description: "检查当前活动 Word 文档结构，返回文档名、路径、段落数、表格数、字数等摘要信息",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
  requiresOfficeApp: "word",
};

const WORD_READ_TEXT_DEF: ToolDefinition = {
  name: "word.readText",
  description:
    "读取当前活动 Word 文档文本。用于理解文档内容、确认编辑结果。maxChars 可限制返回字符数，默认 12000",
  parameters: {
    type: "object",
    properties: {
      maxChars: { type: "number", description: "最多返回字符数，默认 12000" },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
  requiresOfficeApp: "word",
};

const WORD_INSERT_TEXT_DEF: ToolDefinition = {
  name: "word.insertText",
  description: "向当前活动 Word 文档插入文本。position 支持 end/start/selection，默认 end",
  parameters: {
    type: "object",
    properties: {
      text: { type: "string", description: "要插入的文本" },
      position: {
        type: "string",
        enum: ["end", "start", "selection"],
        description: "插入位置，默认 end",
      },
    },
    required: ["text"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

const WORD_INSERT_HEADING_DEF: ToolDefinition = {
  name: "word.insertHeading",
  description: "向当前活动 Word 文档插入标题段落，并套用 Word 标题 1-9 样式。适合新增章节标题",
  parameters: {
    type: "object",
    properties: {
      text: { type: "string", description: "标题文本" },
      level: { type: "integer", minimum: 1, maximum: 9, description: "标题级别 1-9，默认 1" },
      position: {
        type: "string",
        enum: ["end", "start", "selection"],
        description: "插入位置，默认 end",
      },
    },
    required: ["text"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

const WORD_REPLACE_TEXT_DEF: ToolDefinition = {
  name: "word.replaceText",
  description:
    "在当前活动 Word 文档中查找并替换文本，返回替换次数。用于批量修改标题、术语、占位符等",
  parameters: {
    type: "object",
    properties: {
      findText: { type: "string", description: "要查找的文本" },
      replaceText: { type: "string", description: "替换后的文本" },
      matchCase: { type: "boolean", description: "是否区分大小写，默认 false" },
    },
    required: ["findText", "replaceText"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

const WORD_SAVE_DEF: ToolDefinition = {
  name: "word.save",
  description: "保存当前活动 Word 文档。如果指定 saveAsPath 则另存为新文件",
  parameters: {
    type: "object",
    properties: {
      saveAsPath: { type: "string", description: "另存为路径（可选）" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

export const OFFICE_WORD_TOOL_DEFINITIONS: ToolDefinition[] = [
  WORD_OPEN_DEF,
  WORD_INSPECT_DEF,
  WORD_READ_TEXT_DEF,
  WORD_INSERT_TEXT_DEF,
  WORD_INSERT_HEADING_DEF,
  WORD_REPLACE_TEXT_DEF,
  WORD_SAVE_DEF,
];
