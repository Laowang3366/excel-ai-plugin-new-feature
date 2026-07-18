/**
 * Office 工具定义
 *
 * 包含 Word、PowerPoint 和通用 Office 脚本工具。
 */

import type { ToolDefinition } from "../../shared/types";
import {
  APPLY_OPERATIONS,
  INSPECT_OPERATIONS,
  withOfficeOperationDiscriminator,
} from "./officeActionSchemas";
import { OFFICE_RELIABILITY_TOOL_DEFINITIONS } from "./officeReliability";
import { OFFICE_WORD_TOOL_DEFINITIONS } from "./officeWordTools";
import { OFFICE_PRESENTATION_TOOL_DEFINITIONS } from "./officePresentationTools";
import {
  OFFICE_WORKFLOW_STEPS_SCHEMA,
  OFFICE_WORKFLOW_VARIABLES_SCHEMA,
} from "./officeWorkflowSchema";

const OFFICE_CONNECTION_STATUS_DEF: ToolDefinition = {
  name: "office.connection.status",
  description:
    "检测指定 Office 当前窗口环境是否已连接。操作当前窗口、选区或需要 COM 的能力前调用；创建独立磁盘 Excel/Word/PPT 文件可直接使用 office.action.apply 的 Open XML 创建操作，不依赖连接。",
  parameters: {
    type: "object",
    properties: {
      app: {
        type: "string",
        enum: ["excel", "word", "presentation"],
        description: "要检测的 Office 应用类型",
      },
    },
    required: ["app"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_DOCUMENTS_LIST_DEF: ToolDefinition = {
  name: "office.documents.list",
  description:
    "列出所有 Excel、Word、PowerPoint 或 WPS 进程中的已打开文档，返回完整路径、宿主、进程和稳定 instanceId。多窗口操作前先调用本工具定位目标。",
  parameters: {
    type: "object",
    properties: {
      app: {
        type: "string",
        enum: ["excel", "word", "presentation"],
        description: "可选；不填时列出全部 Office 应用",
      },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_DOCUMENTS_ACTIVATE_DEF: ToolDefinition = {
  name: "office.documents.activate",
  description:
    "按 office.documents.list 返回的 instanceId 和完整路径激活指定 Excel、Word、PowerPoint/WPS 文档；仅一个候选时也可按名称或序号。",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用" },
      filePath: { type: "string", description: "目标文档完整路径，优先使用" },
      name: { type: "string", description: "目标文档名称" },
      index: { type: "integer", minimum: 1, description: "目标文档在应用集合中的序号，从 1 开始" },
      instanceId: {
        type: "string",
        description: "office.documents.list 返回的实例标识；多进程或同路径副本必须传",
      },
    },
    required: ["app"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_WORKFLOW_RUN_DEF: ToolDefinition = {
  name: "office.workflow.run",
  description:
    "执行最多 20 个持久化 Office 步骤，支持输出占位符、条件、受控并行、有限重试、超时、取消、崩溃租约恢复和确定性事务撤销重做。",
  parameters: {
    type: "object",
    properties: {
      steps: {
        ...OFFICE_WORKFLOW_STEPS_SCHEMA,
        description:
          "有序步骤；每步结构与 office.action.apply 一致，需提供 app、action、operation、filePath，可选 target/outputPath/preferEngine/params",
      },
      templateId: {
        type: "string",
        description: "使用已保存模板时传模板 ID 或名称，此时可省略 steps",
      },
      variables: OFFICE_WORKFLOW_VARIABLES_SCHEMA,
      workflowId: { type: "string", description: "继续已有流水线时传原 workflowId" },
      resume: { type: "boolean", description: "设为 true 时从 workflowId 记录的失败步骤继续" },
      recoverRunning: {
        type: "boolean",
        description: "确认原执行进程已终止时，接管仍为 running 的工作流",
      },
      leaseMs: {
        type: "integer",
        minimum: 30_000,
        maximum: 1_800_000,
        description: "运行租约时长，30 秒到 30 分钟，默认 5 分钟",
      },
      failureMode: {
        type: "string",
        enum: ["pause", "rollback"],
        description: "失败处理；默认 pause，rollback 会立即整体撤销",
      },
      cancellationMode: {
        type: "string",
        enum: ["pause", "rollback"],
        description: "收到取消请求后的处理，默认 pause",
      },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

const OFFICE_ACTION_INSPECT_DEF: ToolDefinition = {
  name: "office.action.inspect",
  description:
    "统一 Office 高级检查入口。Excel 可检查查询、图表、对象、模板、打印和公式治理；Word/PPT 可检查排版、审阅、母版、动画、备注以及 Excel 链接来源；也可检查 Office 文件结构与事务备份。",
  parameters: withOfficeOperationDiscriminator(
    {
      type: "object",
      properties: {
        app: {
          type: "string",
          enum: ["excel", "word", "presentation"],
          description: "目标应用类型",
        },
        action: {
          type: "string",
          enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"],
          description: "动作类型，默认 inspect",
        },
        operation: {
          type: "string",
          description:
            "检查操作。Word: inspectDocumentFormatting/inspectReferences/inspectRevisions/inspectContentControls/inspectLinkedOfficeContent；Excel: inspectPrintSettings/inspectFormulaDependencies/inspectFormulaBackups/inspectFormulaProtection/inspectPowerQueries/inspectCharts/inspectWorkbookObjects/captureWorkbookTemplate；PowerPoint: inspectPresentationTheme/inspectSlideElements/inspectAnimations/inspectSpeakerNotes/inspectLinkedOfficeContent；通用: inspectFile/layout/tables/listBackups。",
        },
        filePath: { type: "string", description: "Office 文件绝对路径" },
        outputPath: { type: "string", description: "输出文件路径" },
        target: {
          type: "string",
          description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1",
        },
        preferEngine: {
          type: "string",
          enum: ["openxml", "com"],
          description: "首选引擎，默认 openxml",
        },
        params: { type: "object", description: "操作参数" },
      },
      required: ["app", "operation"],
    },
    INSPECT_OPERATIONS,
  ),
  riskLevel: "safe",
  requiresApproval: false,
};

const OFFICE_ACTION_APPLY_DEF: ToolDefinition = {
  name: "office.action.apply",
  description:
    "统一 Office 文件级操作入口，必须提供 filePath。可用 createWorkbook/createDocument/createPresentation 创建不依赖桌面应用连接的新文件，也支持后续编辑、治理、链接报告及原位刷新。复杂多文件任务应交给 office.workflow.run，以获得暂停续跑和整体撤销；当前活动窗口或未保存内容仍使用 range.*、word.*、presentation.*。",
  parameters: withOfficeOperationDiscriminator(
    {
      type: "object",
      properties: {
        app: {
          type: "string",
          enum: ["excel", "word", "presentation"],
          description: "目标应用类型",
        },
        action: {
          type: "string",
          enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"],
          description: "动作类型，必填",
        },
        operation: {
          type: "string",
          description:
            "操作。Excel 新建: createWorkbook；Word 新建: createDocument；PowerPoint 新建与批量页: createPresentation/addSlides；Word 高级: formatLongDocument/manageReferences/manageRevisions/compareDocuments/applyTrackedChanges/mailMerge/batchMailMerge/populateContentControls/manageContentControls/refreshLinkedOfficeContent/relinkLinkedOfficeContent；PowerPoint 高级: insertTable/applyMasterBranding/layoutElements/configureAnimations/configureSlideShow/setSpeakerNotes/exportHandouts/refreshLinkedOfficeContent/relinkLinkedOfficeContent；Excel 跨应用: exportRangeToWord/exportRangeToPresentation/buildReportPackage；通用: snapshot。",
        },
        filePath: { type: "string", description: "Office 文件绝对路径；文件级 apply 必填" },
        outputPath: {
          type: "string",
          description:
            "输出文件路径；修改操作未指定时原地保存并自动备份，导出操作未指定时生成带后缀的新文件",
        },
        target: {
          type: "string",
          description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1",
        },
        preferEngine: {
          type: "string",
          enum: ["openxml", "com"],
          description: "首选引擎，默认 openxml",
        },
        params: {
          type: "object",
          description:
            "参数。Power Query 必须声明 advancedIntent:'refreshable-etl'，创建/更新另需 sourceKind:'external'|'multi-source'；透视表/切片器必须声明 advancedIntent:'interactive-pivot'。Word 修订用 rule，批量合并用 dataSourcePath/outputDirectory/outputFormat；PowerPoint 排版用 edits/align/distribute/crop；Excel 联动用 linked:true、sourceType:range|chart、chartName、linkId，报告包可用 sections。其他高级参数见系统 Office 工具说明。",
        },
      },
      required: ["app", "action", "operation", "filePath"],
    },
    APPLY_OPERATIONS,
  ),
  riskLevel: "moderate",
  requiresApproval: true,
};

const OFFICE_ACTION_VALIDATE_DEF: ToolDefinition = {
  name: "office.action.validate",
  description:
    "统一 Office 只读验证入口。operation 使用 inspectFile、layout、tables 或其他 inspect* 操作；params 可用 containsText、countPath+expectedCount/minCount、outputExists 定义可判定的验证条件。",
  parameters: withOfficeOperationDiscriminator(
    {
      type: "object",
      properties: {
        app: {
          type: "string",
          enum: ["excel", "word", "presentation"],
          description: "目标应用类型",
        },
        action: {
          type: "string",
          enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"],
          description: "动作类型，默认 validate",
        },
        operation: {
          type: "string",
          description:
            "只读检查操作，如 inspectFile、layout、tables、inspectCharts、inspectReferences",
        },
        filePath: { type: "string", description: "Office 文件绝对路径" },
        outputPath: { type: "string", description: "输出文件路径" },
        target: { type: "string", description: "对象定位" },
        preferEngine: {
          type: "string",
          enum: ["openxml", "com"],
          description: "首选引擎，默认 openxml",
        },
        params: {
          type: "object",
          description:
            "验证条件：containsText 字符串或数组；countPath 配合 expectedCount/minCount；outputExists 检查输出文件",
        },
      },
      required: ["app", "operation"],
    },
    INSPECT_OPERATIONS,
  ),
  riskLevel: "safe",
  requiresApproval: false,
};

export const OFFICE_TOOL_DEFINITIONS: ToolDefinition[] = [
  OFFICE_CONNECTION_STATUS_DEF,
  OFFICE_DOCUMENTS_LIST_DEF,
  OFFICE_DOCUMENTS_ACTIVATE_DEF,
  OFFICE_WORKFLOW_RUN_DEF,
  ...OFFICE_RELIABILITY_TOOL_DEFINITIONS,
  ...OFFICE_WORD_TOOL_DEFINITIONS,
  ...OFFICE_PRESENTATION_TOOL_DEFINITIONS,
  OFFICE_ACTION_INSPECT_DEF,
  OFFICE_ACTION_APPLY_DEF,
  OFFICE_ACTION_VALIDATE_DEF,
];
