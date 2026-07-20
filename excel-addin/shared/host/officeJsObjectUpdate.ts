import { mapChartType, toChartTypeLabel } from "./officeJsChartTypes";
import { normalizeSameSheetA1Range } from "./officeJsChartSource";
import { withExcel } from "./officeJsRuntime";
import type {
  ChartInfo,
  ChartUpdateInput,
  HostResult,
  TableInfo,
  TableUpdateInput,
} from "./types";
import { unsupported } from "./types";

function requiredTableUpdateApi(input: TableUpdateInput): "1.3" | "1.13" | null {
  if (input.resizeAddress != null) return "1.13";
  if (
    input.showFilterButton != null ||
    input.showBandedRows != null ||
    input.showBandedColumns != null ||
    input.showFirstColumn != null ||
    input.showLastColumn != null
  ) {
    return "1.3";
  }
  return null;
}

function isExcelApiSupported(version: string): boolean {
  const office = (globalThis as unknown as {
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", version);
  } catch {
    return false;
  }
}

export async function officeJsUpdateTable(
  input: TableUpdateInput,
): Promise<HostResult<TableInfo>> {
  const requiredApi = requiredTableUpdateApi(input);
  if (requiredApi && !isExcelApiSupported(requiredApi)) {
    return unsupported(
      "table.update",
      "office-js",
      `ExcelApi ${requiredApi} is not supported in this host (Office.context.requirements.isSetSupported)`,
      requiredApi === "1.13"
        ? "Table.resize requires ExcelApi 1.13"
        : "Table.showFilterButton/showBandedRows/showBandedColumns/showFirstColumn/showLastColumn require ExcelApi 1.3",
    );
  }
  return withExcel("table.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const table = sheet.tables.getItem(input.tableName);
    if (input.resizeAddress != null) {
      table.resize(
        normalizeSameSheetA1Range(
          input.sheetName,
          input.resizeAddress,
          "resizeAddress",
          "table",
        ),
      );
    }
    if (input.newName != null) table.name = input.newName;
    if (input.style != null) table.style = input.style;
    if (input.showHeaders != null) table.showHeaders = input.showHeaders;
    if (input.showTotals != null) table.showTotals = input.showTotals;
    if (input.showFilterButton != null) table.showFilterButton = input.showFilterButton;
    if (input.showBandedRows != null) table.showBandedRows = input.showBandedRows;
    if (input.showBandedColumns != null) table.showBandedColumns = input.showBandedColumns;
    if (input.showFirstColumn != null) table.showFirstColumn = input.showFirstColumn;
    if (input.showLastColumn != null) table.showLastColumn = input.showLastColumn;
    await context.sync();
    table.load(
      "name,showHeaders,showFilterButton,showTotals,showBandedRows,showBandedColumns,showFirstColumn,showLastColumn,style",
    );
    const range = table.getRange();
    range.load("address");
    await context.sync();
    return {
      name: table.name,
      sheetName: input.sheetName,
      address: range.address,
      hasHeaders: table.showHeaders,
      showFilter: table.showFilterButton,
      showTotals: table.showTotals,
      showBandedRows: table.showBandedRows,
      showBandedColumns: table.showBandedColumns,
      showFirstColumn: table.showFirstColumn,
      showLastColumn: table.showLastColumn,
      style: table.style,
    };
  });
}

export async function officeJsUpdateChart(
  input: ChartUpdateInput,
): Promise<HostResult<ChartInfo>> {
  return withExcel("chart.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName);
    if (input.newName != null) chart.name = input.newName;
    if (input.chartType != null) chart.chartType = mapChartType(input.chartType);
    if (input.title != null) {
      chart.title.text = input.title;
      chart.title.visible = input.showTitle !== false;
    } else if (input.showTitle != null) {
      chart.title.visible = input.showTitle;
    }
    if (input.left != null) chart.left = input.left;
    if (input.top != null) chart.top = input.top;
    if (input.width != null) chart.width = input.width;
    if (input.height != null) chart.height = input.height;
    if (input.style != null) chart.style = input.style;
    if (input.showLegend != null) chart.legend.visible = input.showLegend;
    await context.sync();
    chart.load("name,chartType,style,left,top,width,height");
    chart.title.load("text,visible");
    chart.legend.load("visible");
    await context.sync();
    return {
      name: chart.name,
      sheetName: input.sheetName,
      chartType: toChartTypeLabel(String(chart.chartType)),
      title: chart.title.text,
      titleVisible: chart.title.visible,
      left: chart.left,
      top: chart.top,
      width: chart.width,
      height: chart.height,
      style: chart.style,
      legendVisible: chart.legend.visible,
    };
  });
}
