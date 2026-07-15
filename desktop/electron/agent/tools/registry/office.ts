/**
 * Office 工具定义
 *
 * 包含 Word、PowerPoint 和通用 Office 脚本工具。
 */

import type { ToolDefinition } from "../../shared/types";
import { OFFICE_RELIABILITY_TOOL_DEFINITIONS } from "./officeReliability";
import { OFFICE_WORKFLOW_STEPS_SCHEMA } from "./officeWorkflowSchema";

const OFFICE_CONNECTION_STATUS_DEF: ToolDefinition = {
  name: "office.connection.status",
  description: "检测指定 Office 环境是否已连接。每次执行 Excel/Word/PowerPoint 创建、编辑、读取或当前窗口操作前，先调用本工具判断 connected，再按连接状态选择 office.action.*、range.* 或专用 Office 工具。",
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

const OFFICE_DOCUMENTS_LIST_DEF: ToolDefinition = {
  name: "office.documents.list",
  description: "列出所有 Excel、Word、PowerPoint 或 WPS 进程中的已打开文档，返回完整路径、宿主、进程和稳定 instanceId。多窗口操作前先调用本工具定位目标。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "可选；不填时列出全部 Office 应用" },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_DOCUMENTS_ACTIVATE_DEF: ToolDefinition = {
  name: "office.documents.activate",
  description: "按 office.documents.list 返回的 instanceId 和完整路径激活指定 Excel、Word、PowerPoint/WPS 文档；仅一个候选时也可按名称或序号。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用" },
      filePath: { type: "string", description: "目标文档完整路径，优先使用" },
      name: { type: "string", description: "目标文档名称" },
      index: { type: "integer", minimum: 1, description: "目标文档在应用集合中的序号，从 1 开始" },
      instanceId: { type: "string", description: "office.documents.list 返回的实例标识；多进程或同路径副本必须传" },
    },
    required: ["app"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_WORKFLOW_RUN_DEF: ToolDefinition = {
  name: "office.workflow.run",
  description: "执行最多 20 个持久化 Office 步骤，支持输出占位符、条件、受控并行、有限重试、超时、取消、崩溃租约恢复和确定性事务撤销重做。",
  parameters: {
    type: "object",
    properties: {
      steps: {
        ...OFFICE_WORKFLOW_STEPS_SCHEMA,
        description: "有序步骤；每步结构与 office.action.apply 一致，需提供 app、action、operation、filePath，可选 target/outputPath/preferEngine/params",
      },
      templateId: { type: "string", description: "使用已保存模板时传模板 ID 或名称，此时可省略 steps" },
      variables: { type: "object", description: "模板变量；步骤字符串中的 {{vars.name}} 会在事务快照前展开" },
      workflowId: { type: "string", description: "继续已有流水线时传原 workflowId" },
      resume: { type: "boolean", description: "设为 true 时从 workflowId 记录的失败步骤继续" },
      recoverRunning: { type: "boolean", description: "确认原执行进程已终止时，接管仍为 running 的工作流" },
      leaseMs: { type: "integer", minimum: 30_000, maximum: 1_800_000, description: "运行租约时长，30 秒到 30 分钟，默认 5 分钟" },
      failureMode: { type: "string", enum: ["pause", "rollback"], description: "失败处理；默认 pause，rollback 会立即整体撤销" },
      cancellationMode: { type: "string", enum: ["pause", "rollback"], description: "收到取消请求后的处理，默认 pause" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
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
      level: { type: "integer", minimum: 1, maximum: 9, description: "标题级别 1-9，默认 1" },
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

const OFFICE_ACTION_INSPECT_DEF: ToolDefinition = {
  name: "office.action.inspect",
  description: "统一 Office 高级检查入口。Excel 可检查查询、图表、对象、模板、打印和公式治理；Word/PPT 可检查排版、审阅、母版、动画、备注以及 Excel 链接来源；也可检查 Office 文件结构与事务备份。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，默认 inspect" },
      operation: { type: "string", description: "检查操作。Word: inspectDocumentFormatting/inspectReferences/inspectRevisions/inspectContentControls/inspectLinkedOfficeContent；Excel: inspectPrintSettings/inspectFormulaDependencies/inspectFormulaBackups/inspectFormulaProtection/inspectPowerQueries/inspectCharts/inspectWorkbookObjects/captureWorkbookTemplate；PowerPoint: inspectPresentationTheme/inspectSlideElements/inspectAnimations/inspectSpeakerNotes/inspectLinkedOfficeContent；通用: inspectFile/layout/tables/listBackups。" },
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
  description: "统一 Office 文件级高级操作入口，必须提供 filePath。支持 Excel 治理、Word/PPT 高级编辑、Excel 链接报告及原位刷新。复杂多文件任务应交给 office.workflow.run，以获得暂停续跑和整体撤销；当前活动窗口或未保存内容仍使用 range.*、word.*、presentation.*。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，必填" },
      operation: { type: "string", description: "操作。Word: formatLongDocument/manageReferences/manageRevisions/compareDocuments/applyTrackedChanges/mailMerge/batchMailMerge/populateContentControls/manageContentControls/refreshLinkedOfficeContent/relinkLinkedOfficeContent；PowerPoint: insertTable/applyMasterBranding/layoutElements/configureAnimations/configureSlideShow/setSpeakerNotes/exportHandouts/refreshLinkedOfficeContent/relinkLinkedOfficeContent；Excel 跨应用: exportRangeToWord/exportRangeToPresentation/buildReportPackage（updateExisting 按 linkId 增量维护）；通用: snapshot。" },
      filePath: { type: "string", description: "Office 文件绝对路径；文件级 apply 必填" },
      outputPath: { type: "string", description: "输出文件路径；修改操作未指定时原地保存并自动备份，导出操作未指定时生成带后缀的新文件" },
      target: { type: "string", description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "参数。Power Query 必须声明 advancedIntent:'refreshable-etl'，创建/更新另需 sourceKind:'external'|'multi-source'；透视表/切片器必须声明 advancedIntent:'interactive-pivot'。Word 修订用 rule，批量合并用 dataSourcePath/outputDirectory/outputFormat；PowerPoint 排版用 edits/align/distribute/crop；Excel 联动用 linked:true、sourceType:range|chart、chartName、linkId，报告包可用 sections。其他高级参数见系统 Office 工具说明。" },
    },
    required: ["app", "action", "operation", "filePath"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

const OFFICE_ACTION_VALIDATE_DEF: ToolDefinition = {
  name: "office.action.validate",
  description: "统一 Office 只读验证入口。operation 使用 inspectFile、layout、tables 或其他 inspect* 操作；params 可用 containsText、countPath+expectedCount/minCount、outputExists 定义可判定的验证条件。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型，默认 validate" },
      operation: { type: "string", description: "只读检查操作，如 inspectFile、layout、tables、inspectCharts、inspectReferences" },
      filePath: { type: "string", description: "Office 文件绝对路径" },
      outputPath: { type: "string", description: "输出文件路径" },
      target: { type: "string", description: "对象定位" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "验证条件：containsText 字符串或数组；countPath 配合 expectedCount/minCount；outputExists 检查输出文件" },
    },
    required: ["app", "operation"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const OFFICE_TOOL_DEFINITIONS: ToolDefinition[] = [
  OFFICE_CONNECTION_STATUS_DEF,
  OFFICE_DOCUMENTS_LIST_DEF,
  OFFICE_DOCUMENTS_ACTIVATE_DEF,
  OFFICE_WORKFLOW_RUN_DEF,
  ...OFFICE_RELIABILITY_TOOL_DEFINITIONS,
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
];
