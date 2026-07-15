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
  resolveOfficeAdvancedIntents,
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
      .map(({ key, content }) => ({
        key,
        content: key === "office-tools"
          ? renderPromptTemplate(content, {
            ADVANCED_EXCEL_BOUNDARY: buildAdvancedExcelBoundary(context),
          })
          : content,
      })),
  );
}

function buildAdvancedExcelBoundary(context: PromptBuildContext): string {
  const intents = resolveOfficeAdvancedIntents(context);
  const rules = [
    "Excel：值、公式、格式、固定汇总用 `range.write`，数据量不是升级理由。禁止为写值创建高级查询或交互透视对象。",
  ];
  if (intents.has("refreshable-etl")) {
    rules.push(
      "本轮明确要求外部/多来源可刷新 ETL，可开放 `createPowerQuery/managePowerQuery`；须 `filePath`、`params.advancedIntent:\"refreshable-etl\"`，创建/更新另传 `sourceKind:\"external\"|\"multi-source\"`。",
    );
  }
  if (intents.has("interactive-pivot")) {
    rules.push(
      "本轮明确要求交互式透视，可开放 `createPivotTable/refreshPivotTables/addSlicer`；须 `params.advancedIntent:\"interactive-pivot\"`，创建时明确源区域和字段。",
    );
  }
  if (rules.length === 1) {
    rules.push("本轮未检测到上述高级意图，相关 operation 不向模型开放。");
  }
  return `- ${rules.join(" ")}`;
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
