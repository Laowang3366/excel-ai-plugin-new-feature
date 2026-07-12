import basePrompt from "./templates/system/base.zh-CN.md?raw";
import securityPrompt from "./templates/system/security.zh-CN.md?raw";
import formulaPrompt from "./templates/scenarios/formula.zh-CN.md?raw";
import ocrInvoicePrompt from "./templates/scenarios/ocr-invoice.zh-CN.md?raw";
import officeToolsPrompt from "./templates/scenarios/office-tools.zh-CN.md?raw";
import generalOfficePrompt from "./templates/scenarios/general-office.zh-CN.md?raw";
import macroPrompt from "./templates/scenarios/macro.zh-CN.md?raw";
import runtimeEnvironmentPrompt from "./templates/runtime/environment.zh-CN.md?raw";
import dynamicArrayEnabledPrompt from "./templates/runtime/dynamic-array-enabled.zh-CN.md?raw";
import dynamicArrayDisabledPrompt from "./templates/runtime/dynamic-array-disabled.zh-CN.md?raw";
import { composePromptSections, renderPromptTemplate } from "./promptComposer";
import {
  resolvePromptScenarios,
  type PromptRoutingContext,
  type PromptScenario,
} from "./promptRouting";

export { appendFolderContext } from "./sections/folderContextPrompt";
export type { FolderFileItem } from "./sections/folderContextPrompt";

export interface PromptBuildContext extends PromptRoutingContext {
  folderId?: string;
}

export interface RuntimePromptContext {
  officeConnectionStatus: string;
  dynamicArrayFunctionsEnabled: boolean;
  now?: Date;
}

interface ContextualPromptDefinition {
  key: string;
  content: string;
  scenario: PromptScenario;
}

const baseSections = [
  { key: "base", content: basePrompt },
  { key: "security", content: securityPrompt },
];

const contextualSections: ContextualPromptDefinition[] = [
  { key: "formula", content: formulaPrompt, scenario: "formula" },
  { key: "ocr-invoice", content: ocrInvoicePrompt, scenario: "ocr-invoice" },
  { key: "office-tools", content: officeToolsPrompt, scenario: "office-tools" },
  { key: "general-office", content: generalOfficePrompt, scenario: "general-office" },
  { key: "macro", content: macroPrompt, scenario: "macro" },
];

export function buildSystemPrompt(): string {
  return composePromptSections(baseSections);
}

export function buildContextualPromptSections(context: PromptBuildContext = {}): string {
  const scenarios = resolvePromptScenarios(context);
  return composePromptSections(
    contextualSections
      .filter((section) => scenarios.has(section.scenario))
      .map(({ key, content }) => ({ key, content })),
  );
}

export function buildRuntimePromptSection(context: RuntimePromptContext): string {
  const now = context.now ?? new Date();
  const dateFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "long",
  });
  const timeFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  const dynamicArraySupport = context.dynamicArrayFunctionsEnabled
    ? dynamicArrayEnabledPrompt
    : dynamicArrayDisabledPrompt;
  return renderPromptTemplate(runtimeEnvironmentPrompt, {
    OFFICE_CONNECTION_CONTEXT: `- Office 应用连接状态：${context.officeConnectionStatus}`,
    DYNAMIC_ARRAY_SUPPORT: dynamicArraySupport,
    CURRENT_DATE: dateFormatter.format(now),
    CURRENT_TIME: timeFormatter.format(now),
  });
}
