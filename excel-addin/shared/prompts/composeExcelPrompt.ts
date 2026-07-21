import {
  buildAddInHardBoundary,
  buildAdvancedExcelBoundary,
} from "./advancedExcelBoundary";
import { getPromptText } from "./loadPrompts";
import {
  appendPromptSections,
  composePromptSections,
  renderPromptTemplate,
  type PromptSection,
} from "./promptComposer";
import {
  resolveExcelPromptScenarios,
  type PromptRoutingContext,
} from "./promptRouting";
import {
  formatOfficeConnectionContext,
  formatRuntimeDateTime,
} from "./runtimeContext";

const SCENARIO_TEMPLATE: Record<string, string> = {
  formula: "scenarios/formula.zh-CN.md",
  "office-tools": "scenarios/office-tools.zh-CN.md",
  "general-office": "scenarios/general-office.zh-CN.md",
  macro: "scenarios/macro.zh-CN.md",
  "ocr-invoice": "scenarios/ocr-invoice.zh-CN.md",
};

export interface ExcelPromptBuildOptions {
  routing: PromptRoutingContext;
  /** Raw connection status string, e.g. "connected (office-js)" */
  officeConnectionStatus?: string;
  runtimeVars?: Record<string, string>;
  dynamicArrayEnabled?: boolean;
  now?: Date;
}

/** Build Excel-only system + scenario + runtime prompt from synced templates. */
export function composeExcelSystemPrompt(options: ExcelPromptBuildOptions): string {
  const base = composePromptSections([
    { key: "system/base", content: getPromptText("system/base.zh-CN.md") },
    { key: "system/security", content: getPromptText("system/security.zh-CN.md") },
  ]);

  const scenarios = resolveExcelPromptScenarios(options.routing);
  const scenarioSections: PromptSection[] = [];
  for (const scenario of scenarios) {
    const templateId = SCENARIO_TEMPLATE[scenario];
    if (!templateId) continue;
    let content = getPromptText(templateId);
    if (scenario === "office-tools" && content.includes("{{ADVANCED_EXCEL_BOUNDARY}}")) {
      content = renderPromptTemplate(content, {
        ADVANCED_EXCEL_BOUNDARY:
          options.runtimeVars?.ADVANCED_EXCEL_BOUNDARY ??
          buildAdvancedExcelBoundary(options.routing),
      });
    }
    scenarioSections.push({ key: `scenario/${scenario}`, content });
  }

  const dynamicArrayId = options.dynamicArrayEnabled
    ? "runtime/dynamic-array-enabled.zh-CN.md"
    : "runtime/dynamic-array-disabled.zh-CN.md";
  const envTemplate = getPromptText("runtime/environment.zh-CN.md");
  const clock = formatRuntimeDateTime(options.now ?? new Date());
  const runtimeVars = {
    OFFICE_CONNECTION_CONTEXT: formatOfficeConnectionContext(
      options.officeConnectionStatus ??
        options.runtimeVars?.OFFICE_CONNECTION_STATUS ??
        "unknown",
    ),
    DYNAMIC_ARRAY_SUPPORT: getPromptText(dynamicArrayId),
    CURRENT_DATE: clock.CURRENT_DATE,
    CURRENT_TIME: clock.CURRENT_TIME,
    ...options.runtimeVars,
  };
  // Ensure connection context uses desktop semantics unless fully overridden.
  if (!options.runtimeVars?.OFFICE_CONNECTION_CONTEXT) {
    runtimeVars.OFFICE_CONNECTION_CONTEXT = formatOfficeConnectionContext(
      options.officeConnectionStatus ?? "unknown",
    );
  }
  if (!options.runtimeVars?.CURRENT_DATE) {
    runtimeVars.CURRENT_DATE = clock.CURRENT_DATE;
  }
  if (!options.runtimeVars?.CURRENT_TIME) {
    runtimeVars.CURRENT_TIME = clock.CURRENT_TIME;
  }

  const runtimeSection: PromptSection = {
    key: "runtime/environment",
    content: renderPromptTemplate(envTemplate, runtimeVars),
  };

  // Always last: overrides synced desktop macro/Open XML/COM narratives for all routes.
  const hardBoundary: PromptSection = {
    key: "addin/hard-boundary",
    content: buildAddInHardBoundary(),
  };

  return appendPromptSections(base, [...scenarioSections, runtimeSection, hardBoundary]);
}
