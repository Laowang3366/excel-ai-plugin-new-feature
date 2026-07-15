import type { ToolDefinition } from "../../shared/types";
import { OFFICE_WORKFLOW_STEPS_SCHEMA } from "./officeWorkflowSchema";

const APP_PROPERTY = {
  type: "string",
  enum: ["excel", "word", "presentation"],
  description: "Office 应用类型",
} as const;

export const OFFICE_RELIABILITY_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "office.workflow.status",
    description:
      "查看持久化办公流水线。传 workflowId 返回步骤、产物、失败位置和下一步；不传时列出全部流水线。",
    parameters: {
      type: "object",
      properties: {
        workflowId: { type: "string", description: "工作流 ID，可从 office.workflow.run 结果取得" },
      },
      required: [],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "office.workflow.cancel",
    description:
      "请求在当前 Office 步骤结束后的安全边界取消工作流。记录保留，可随后恢复；不会在文件写入中途强杀进程。",
    parameters: {
      type: "object",
      properties: { workflowId: { type: "string", description: "正在运行或暂停的工作流 ID" } },
      required: ["workflowId"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
  {
    name: "office.workflow.template.list",
    description: "列出已保存的 Office 流水线模板及其稳定 ID、说明和步骤。",
    parameters: { type: "object", properties: {}, required: [] },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "office.workflow.template.save",
    description:
      "保存或更新可复用 Office 流水线模板。路径和参数可使用 {{vars.name}}，运行时由 variables 提供。",
    parameters: {
      type: "object",
      properties: {
        templateId: { type: "string", description: "更新已有模板时传 ID" },
        name: { type: "string", description: "模板名称" },
        description: { type: "string" },
        steps: {
          ...OFFICE_WORKFLOW_STEPS_SCHEMA,
          description: "与 office.workflow.run.steps 相同",
        },
      },
      required: ["name", "steps"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
  {
    name: "office.workflow.template.delete",
    description: "删除指定 Office 流水线模板，不影响历史运行和事务记录。",
    parameters: {
      type: "object",
      properties: { templateId: { type: "string" } },
      required: ["templateId"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
  {
    name: "office.objects.list",
    description: "按实例和完整路径列出已打开 Office 文档中的可选对象，返回可稳定复用的 locator。",
    parameters: {
      type: "object",
      properties: {
        app: APP_PROPERTY,
        filePath: {
          type: "string",
          description: "已打开文档的完整路径，必须来自 office.documents.list",
        },
        instanceId: {
          type: "string",
          description: "office.documents.list 返回的实例标识；存在多个候选时必须传",
        },
        kind: {
          type: "string",
          description: "可选对象类型筛选，如 sheet/chart/page/slide/shape/table",
        },
      },
      required: ["app", "filePath"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "office.objects.activate",
    description:
      "按完整文件路径和 office.objects.list 返回的 locator 激活工作表、页面、幻灯片或具体对象，避免操作同名文件和后台窗口。",
    parameters: {
      type: "object",
      properties: {
        app: APP_PROPERTY,
        filePath: { type: "string", description: "目标文档完整路径" },
        instanceId: {
          type: "string",
          description: "office.documents.list 返回的实例标识；存在多个候选时必须传",
        },
        locator: { type: "string", description: "office.objects.list 返回的原始 locator" },
      },
      required: ["app", "filePath", "locator"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
  {
    name: "office.transaction.list",
    description: "列出办公流水线产生的事务，包含状态、步骤、产物和修改清单。",
    parameters: { type: "object", properties: {}, required: [] },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "office.transaction.inspect",
    description: "查看指定 Office 事务的步骤、文件快照、产物和逐项修改清单。",
    parameters: {
      type: "object",
      properties: { transactionId: { type: "string", description: "事务 ID" } },
      required: ["transactionId"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "office.transaction.undo",
    description:
      "整体撤销 Office 事务。默认先校验文件未被事务外修改；发生冲突时返回清单且不覆盖，用户明确确认后才可 force。",
    parameters: {
      type: "object",
      properties: {
        transactionId: { type: "string", description: "事务 ID" },
        force: { type: "boolean", description: "确认丢弃事务后的外部修改时设为 true" },
      },
      required: ["transactionId"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
  {
    name: "office.transaction.redo",
    description:
      "从确定的 after 快照重做已撤销事务，不重新调用外部工具。默认校验文件未被事务外修改，冲突后需用户确认 force。",
    parameters: {
      type: "object",
      properties: {
        transactionId: { type: "string", description: "已撤销事务的 ID" },
        force: { type: "boolean", description: "确认覆盖撤销后的外部修改时设为 true" },
      },
      required: ["transactionId"],
    },
    riskLevel: "moderate",
    requiresApproval: true,
  },
];
