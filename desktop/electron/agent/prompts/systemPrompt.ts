/**
 * 系统提示词组装入口。
 *
 * 提示词正文按业务入口拆到 sections：
 * - modelPrompt: 模型基础行为、工作流程、附件处理和最终回复格式。
 * - formulaAssistantPrompt: 公式助手模块专用规则。
 * - officeToolsPrompt: Excel、Word、PowerPoint、Shell 和脚本工具选择。
 * - permissionPrompt/scriptPrompt/qualityPrompt/scenarioPrompt: 权限、脚本规范、质量守则和通用场景。
 * - folderContextPrompt: 当前工作文件夹上下文追加。
 */

import { roleAndWorkflow } from "./sections/modelPrompt";
import { scenarioFormula } from "./sections/formulaAssistantPrompt";
import { toolSelectionGuide } from "./sections/officeToolsPrompt";
import { permissionRules } from "./sections/permissionPrompt";
import { scriptSpec } from "./sections/scriptPrompt";
import { qualityGuardrails } from "./sections/qualityPrompt";
import { buildScenarioPromptSection } from "./sections/scenarioPrompt";

export { appendFolderContext } from "./sections/folderContextPrompt";
export type { FolderFileItem } from "./sections/folderContextPrompt";

export function buildSystemPrompt(): string {
  return [
    roleAndWorkflow(),
    toolSelectionGuide(),
    permissionRules(),
    scriptSpec(),
    qualityGuardrails(),
    [
      "## 场景化操作指南",
      "",
      scenarioFormula(),
      buildScenarioPromptSection(),
    ].join("\n"),
  ].join("\n\n");
}
