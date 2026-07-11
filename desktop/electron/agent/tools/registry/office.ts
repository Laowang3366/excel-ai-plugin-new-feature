/**
 * Office 工具定义
 *
 * 包含 Word、PowerPoint 和通用 Office 脚本工具。
 */

import type { ToolDefinition } from "../../shared/types";

const OFFICE_CONNECTION_STATUS_DEF: ToolDefinition = {
  name: "office.connection.status",
  description: "检测指定 Office 环境是否已连接。每次执行 Excel/Word/PowerPoint 创建、编辑、读取或当前窗口操作前，先调用本工具判断 connected，再按连接状态选择 office.action.*、range.*、python.execute 或专用 COM 工具。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "要检测的 Office 应用类型" },
    },
    required: ["app"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const WORD_OPEN_DEF: ToolDefinition = {
  name: "word.open",
  description: "把指定路径的 Word 文档（.doc/.docx）打开到 Word 应用窗口并设为活动文档。文件级读取/编辑优先使用 office.action.*",
  parameters: {
    type: "object",
    properties: {
      filePath: { type: "string", description: "Word 文档的绝对路径，如 C:\\Users\\用户\\Desktop\\报告.docx" },
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
  description: "读取当前活动 Word 文档文本。用于理解文档内容、确认编辑结果。maxChars 可限制返回字符数，默认 12000",
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
      position: { type: "string", enum: ["end", "start", "selection"], description: "插入位置，默认 end" },
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
      level: { type: "number", description: "标题级别 1-9，默认 1" },
      position: { type: "string", enum: ["end", "start", "selection"], description: "插入位置，默认 end" },
    },
    required: ["text"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "word",
};

const WORD_REPLACE_TEXT_DEF: ToolDefinition = {
  name: "word.replaceText",
  description: "在当前活动 Word 文档中查找并替换文本，返回替换次数。用于批量修改标题、术语、占位符等",
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

const PRESENTATION_OPEN_DEF: ToolDefinition = {
  name: "presentation.open",
  description: "把指定路径的 PowerPoint 演示文稿（.ppt/.pptx）打开到 PowerPoint 应用窗口并设为活动演示文稿。文件级读取/编辑优先使用 office.action.*",
  parameters: {
    type: "object",
    properties: {
      filePath: { type: "string", description: "演示文稿的绝对路径，如 C:\\Users\\用户\\Desktop\\汇报.pptx" },
    },
    required: ["filePath"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_INSPECT_DEF: ToolDefinition = {
  name: "presentation.inspect",
  description: "检查当前活动 PowerPoint 演示文稿结构，返回文件名、路径、幻灯片数及每页文本形状摘要",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_READ_SLIDE_DEF: ToolDefinition = {
  name: "presentation.readSlide",
  description: "读取指定幻灯片的文本内容和文本形状列表。slideIndex 从 1 开始",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "幻灯片序号，从 1 开始" },
    },
    required: ["slideIndex"],
  },
  riskLevel: "safe",
  requiresApproval: false,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_ADD_SLIDE_DEF: ToolDefinition = {
  name: "presentation.addSlide",
  description: "在当前活动演示文稿末尾添加幻灯片，可同时写入标题和正文",
  parameters: {
    type: "object",
    properties: {
      title: { type: "string", description: "幻灯片标题（可选）" },
      body: { type: "string", description: "幻灯片正文（可选）" },
      layout: { type: "string", enum: ["title", "title_body", "blank"], description: "版式，默认 title_body" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_SET_SHAPE_TEXT_DEF: ToolDefinition = {
  name: "presentation.setShapeText",
  description: "设置指定幻灯片中文本形状的文字。可用 shapeName 或 shapeIndex 指定目标，未指定时写入第一个文本形状",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "幻灯片序号，从 1 开始" },
      text: { type: "string", description: "要写入的文本" },
      shapeName: { type: "string", description: "形状名称（可选）" },
      shapeIndex: { type: "number", description: "形状序号，从 1 开始（可选）" },
    },
    required: ["slideIndex", "text"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_REPLACE_TEXT_DEF: ToolDefinition = {
  name: "presentation.replaceText",
  description: "在当前 PowerPoint 演示文稿全部文本形状中查找并替换文本，返回替换次数",
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
  requiresOfficeApp: "presentation",
};

const PRESENTATION_SAVE_DEF: ToolDefinition = {
  name: "presentation.save",
  description: "保存当前活动 PowerPoint 演示文稿。如果指定 saveAsPath 则另存为新文件",
  parameters: {
    type: "object",
    properties: {
      saveAsPath: { type: "string", description: "另存为路径（可选）" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "presentation",
};

const OFFICE_ACTION_INSPECT_DEF: ToolDefinition = {
  name: "office.action.inspect",
  description: "统一 Office 高级检查入口。用于检查 Excel/Word/PPT 的结构、对象、样式、表格、图表和图片信息",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，默认 inspect" },
      operation: { type: "string", description: "具体检查操作，如 inspectFile、layout、tables" },
      filePath: { type: "string", description: "Office 文件绝对路径" },
      outputPath: { type: "string", description: "输出文件路径" },
      target: { type: "string", description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "操作参数" },
    },
    required: ["app", "operation"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_ACTION_APPLY_DEF: ToolDefinition = {
  name: "office.action.apply",
  description: "统一 Office 高级操作入口。优先使用项目内置 Open XML 处理 Excel/Word/PPT 文件级创建和编辑，不依赖 openpyxl、python-docx、python-pptx 或现场 pip。Excel 文件级常用 operation: createWorkbook、writeRange、setDataValidation、applyConditionalFormatting、styleTable、insertChart；Word 文件级常用 operation: createDocument、replaceText、applyHeadingStyles、styleTables、setHeaderFooter；PPT 文件级常用 operation: createPresentation、deleteSlides、applyTheme。已连接 Office 且要操作当前窗口时才优先使用 range.*、python.execute 或专用 COM 工具。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，必填" },
      operation: { type: "string", description: "具体操作，如 createWorkbook、writeRange、createDocument、createPresentation、deleteSlides、replaceText、styleTable、snapshot、setDataValidation、applyHeadingStyles、applyTheme、insertChart" },
      filePath: { type: "string", description: "Office 文件绝对路径" },
      outputPath: { type: "string", description: "输出文件路径；未指定时由实现生成副本" },
      target: { type: "string", description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "操作参数" },
    },
    required: ["app", "action", "operation"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

const OFFICE_ACTION_VALIDATE_DEF: ToolDefinition = {
  name: "office.action.validate",
  description: "统一 Office 高级操作验证入口。用于验证对象存在、输出文件生成、样式或数量变化",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，默认 validate" },
      operation: { type: "string", description: "具体验证操作，如 inspectFile、replaceText、styleTable" },
      filePath: { type: "string", description: "Office 文件绝对路径" },
      outputPath: { type: "string", description: "输出文件路径" },
      target: { type: "string", description: "对象定位" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "验证参数" },
    },
    required: ["app", "operation"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_SCRIPT_EXECUTE_DEF: ToolDefinition = {
  name: "office.script.execute",
  description: "在 Word 或 PowerPoint COM 对象上执行 PowerShell 脚本，用于完成专用工具覆盖不到的复杂编辑。脚本中会自动注入 $app 变量",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["word", "presentation"], description: "目标 Office 应用" },
      code: { type: "string", description: "PowerShell 脚本片段，自动可用 $app 变量" },
    },
    required: ["app", "code"],
  },
  riskLevel: "dangerous",
  requiresApproval: true,
  requiresOfficeApp: "any",
};

export const OFFICE_TOOL_DEFINITIONS: ToolDefinition[] = [
  OFFICE_CONNECTION_STATUS_DEF,
  WORD_OPEN_DEF,
  WORD_INSPECT_DEF,
  WORD_READ_TEXT_DEF,
  WORD_INSERT_TEXT_DEF,
  WORD_INSERT_HEADING_DEF,
  WORD_REPLACE_TEXT_DEF,
  WORD_SAVE_DEF,
  PRESENTATION_OPEN_DEF,
  PRESENTATION_INSPECT_DEF,
  PRESENTATION_READ_SLIDE_DEF,
  PRESENTATION_ADD_SLIDE_DEF,
  PRESENTATION_SET_SHAPE_TEXT_DEF,
  PRESENTATION_REPLACE_TEXT_DEF,
  PRESENTATION_SAVE_DEF,
  OFFICE_ACTION_INSPECT_DEF,
  OFFICE_ACTION_APPLY_DEF,
  OFFICE_ACTION_VALIDATE_DEF,
  OFFICE_SCRIPT_EXECUTE_DEF,
];
