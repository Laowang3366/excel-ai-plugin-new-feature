import type { ToolDefinition } from "../../shared/types";

const PRESENTATION_OPEN_DEF: ToolDefinition = {
  name: "presentation.open",
  description:
    "把指定路径的 PowerPoint 演示文稿（.ppt/.pptx）打开到 PowerPoint 应用窗口并设为活动演示文稿。文件级读取/编辑优先使用 office.action.*",
  parameters: {
    type: "object",
    properties: {
      filePath: {
        type: "string",
        description: "演示文稿的绝对路径，如 C:\\Users\\用户\\Desktop\\汇报.pptx",
      },
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
      slideIndex: { type: "integer", minimum: 1, description: "幻灯片序号，从 1 开始" },
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
      layout: {
        type: "string",
        enum: ["title", "title_body", "blank"],
        description: "版式，默认 title_body",
      },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  requiresOfficeApp: "presentation",
};

const PRESENTATION_SET_SHAPE_TEXT_DEF: ToolDefinition = {
  name: "presentation.setShapeText",
  description:
    "设置指定幻灯片中文本形状的文字。可用 shapeName 或 shapeIndex 指定目标，未指定时写入第一个文本形状",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "integer", minimum: 1, description: "幻灯片序号，从 1 开始" },
      text: { type: "string", description: "要写入的文本" },
      shapeName: { type: "string", description: "形状名称（可选）" },
      shapeIndex: { type: "integer", minimum: 1, description: "形状序号，从 1 开始（可选）" },
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

export const OFFICE_PRESENTATION_TOOL_DEFINITIONS: ToolDefinition[] = [
  PRESENTATION_OPEN_DEF,
  PRESENTATION_INSPECT_DEF,
  PRESENTATION_READ_SLIDE_DEF,
  PRESENTATION_ADD_SLIDE_DEF,
  PRESENTATION_SET_SHAPE_TEXT_DEF,
  PRESENTATION_REPLACE_TEXT_DEF,
  PRESENTATION_SAVE_DEF,
];
