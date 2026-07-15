import type { OfficeActionInput, OfficeActionKind } from "./types";

export const SAFE_ACTION_OPERATIONS = new Set([
  "inspectFile",
  "layout",
  "tables",
  "listBackups",
  "traceFormulaDependencies",
  "inspectFormulaDependencies",
  "inspectFormulaBackups",
  "inspectFormulaProtection",
  "inspectPrintSettings",
  "inspectDocumentFormatting",
  "inspectReferences",
  "inspectRevisions",
  "inspectContentControls",
  "inspectPowerQueries",
  "inspectCharts",
  "inspectWorkbookObjects",
  "captureWorkbookTemplate",
  "inspectWorkbookFormatting",
  "inspectPresentationTheme",
  "inspectSlideElements",
  "inspectAnimations",
  "inspectSpeakerNotes",
  "inspectLinkedOfficeContent",
]);

export function officeActionOperationError(
  action: OfficeActionKind,
  operation: string,
): string | undefined {
  if ((action === "inspect" || action === "validate") && !SAFE_ACTION_OPERATIONS.has(operation)) {
    return `${action} 仅允许只读 Office 操作；修改文件请使用 office.action.apply`;
  }
  return undefined;
}

export function officeAdvancedOperationError(
  input: Pick<OfficeActionInput, "app" | "operation" | "target" | "params">,
): string | undefined {
  if (input.app !== "excel") return undefined;
  const params = input.params || {};

  if (input.operation === "createPowerQuery" || input.operation === "managePowerQuery") {
    if (params.advancedIntent !== "refreshable-etl") {
      return "Power Query 仅允许明确的外部/多来源可刷新 ETL；params.advancedIntent 必须为 refreshable-etl";
    }
    if (!nonEmptyString(params.name)) return "Power Query 操作需要明确的 params.name";
    const command =
      input.operation === "createPowerQuery"
        ? "upsert"
        : nonEmptyString(params.command) || "upsert";
    if (["create", "update", "upsert"].includes(command)) {
      if (params.sourceKind !== "external" && params.sourceKind !== "multi-source") {
        return "创建或更新 Power Query 时 params.sourceKind 必须为 external 或 multi-source";
      }
      if (!nonEmptyString(params.mFormula)) return "创建或更新 Power Query 需要 params.mFormula";
      const loadError = powerQueryLoadError(params, true);
      if (loadError) return loadError;
    }
    if (command === "load") {
      const loadError = powerQueryLoadError(params, true);
      if (loadError) return loadError;
    }
    if (["duplicate", "rename"].includes(command) && !nonEmptyString(params.newName)) {
      return `Power Query ${command} 需要 params.newName`;
    }
    return undefined;
  }

  if (["createPivotTable", "refreshPivotTables", "addSlicer"].includes(input.operation)) {
    if (params.advancedIntent !== "interactive-pivot") {
      return "透视表和切片器仅允许明确的交互式透视需求；params.advancedIntent 必须为 interactive-pivot";
    }
    if (input.operation === "createPivotTable") {
      if (!input.target?.startsWith("range:") || !input.target.slice("range:".length).trim()) {
        return "创建透视表需要明确的 range: 源区域 target";
      }
      if (
        !["rowFields", "columnFields", "filterFields", "dataFields"].some((key) =>
          hasPivotField(params[key]),
        )
      ) {
        return "创建透视表至少需要一个行、列、筛选或数据字段";
      }
    }
    if (input.operation === "addSlicer") {
      if (!nonEmptyString(params.pivotName)) return "添加切片器需要明确的 params.pivotName";
      if (!nonEmptyString(params.field)) return "添加切片器需要 params.field";
    }
  }
  return undefined;
}

function powerQueryLoadError(
  params: Record<string, unknown>,
  requireLoadMode: boolean,
): string | undefined {
  const loadMode = nonEmptyString(params.loadMode);
  if (!loadMode && !requireLoadMode) return undefined;
  if (!loadMode || !["worksheet", "dataModel", "connectionOnly"].includes(loadMode)) {
    return "Power Query 加载需要 params.loadMode 为 worksheet、dataModel 或 connectionOnly";
  }
  if (loadMode === "worksheet" && !nonEmptyString(params.destination)) {
    return "Power Query 工作表加载需要明确的 params.destination";
  }
  return undefined;
}

function nonEmptyString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function hasPivotField(value: unknown): boolean {
  return (
    Array.isArray(value) &&
    value.some((field) => {
      if (nonEmptyString(field)) return true;
      return (
        typeof field === "object" &&
        field !== null &&
        nonEmptyString((field as Record<string, unknown>).name) !== ""
      );
    })
  );
}
