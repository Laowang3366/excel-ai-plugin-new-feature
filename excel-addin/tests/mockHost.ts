import type {
  CellValue,
  ChartInfo,
  ChartSeriesInfo,
  ChartType,
  HostAdapter,
  HostResult,
  HostStatus,
  RangeData,
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeFormat,
  RangeFormatData,
  RangeInsertInput,
  RangeMutationInfo,
  SelectionInfo,
  SheetInfo,
  TableInfo,
  TableUpdateInput,
  TableUnlistInfo,
  WorkbookInspectInfo,
} from "../shared/host/types";
import type {
  TableFilterApplyInput,
  TableFilterClearInput,
  TableFilterGetInput,
  TableFilterInfo,
} from "../shared/host/tableFilterTypes";
import type {
  TableSortApplyInput,
  TableSortClearInput,
  TableSortGetInput,
  TableSortInfo,
} from "../shared/host/tableSortTypes";
import type {
  FormulaProtectionInspectInfo,
  FormulaProtectionInspectInput,
  FormulaProtectionManageInfo,
  FormulaProtectionManageInput,
} from "../shared/host/formulaProtectionTypes";
import type {
  ChartSeriesAddInput,
  ChartSeriesAddResult,
  ChartSeriesDeleteResult,
  ChartSeriesUpdateInput,
} from "../shared/host/chartSeriesTypes";
import type {
  ChartSeriesValuesInfo,
  ChartSeriesValuesUpdateInput,
} from "../shared/host/chartSeriesValuesTypes";
import type {
  ChartSeriesBubbleSizesInfo,
  ChartSeriesBubbleSizesUpdateInput,
} from "../shared/host/chartSeriesBubbleSizesTypes";
import type { ChartImageGetInput, ChartImageInfo } from "../shared/host/chartImageTypes";
import type { RangeImageGetInput, RangeImageInfo } from "../shared/host/rangeImageTypes";
import type { ChartSourceUpdateInput } from "../shared/host/chartSourceTypes";
import type { ChartAxisUpdateInput } from "../shared/host/chartAxisTypes";
import type { ChartAxisInfo } from "../shared/host/chartAxisTypes";
import type {
  ChartDataLabelsInfo,
  ChartDataLabelsUpdateInput,
} from "../shared/host/chartDataLabelsTypes";
import type {
  ChartSeriesAxisGroupInfo,
  ChartSeriesAxisGroupUpdateInput,
} from "../shared/host/chartSeriesAxisGroupTypes";
import {
  normalizeSameSheetA1Range,
  normalizeSameSheetSourceRange,
} from "../shared/host/officeJsChartSource";
import { createMockStructureState } from "./mockStructure";
import { ok, unsupported } from "../shared/host/types";

function key(sheet: string, address: string): string {
  return `${sheet}!${address.toUpperCase()}`;
}

function chartKey(sheetName: string, chartName: string): string {
  return `${sheetName}\0${chartName}`;
}

type MockSeries = {
  name: string;
  chartType: ChartType | string;
  smooth: boolean;
  /** Persisted host bubble-size source string for MockHost parity. */
  bubbleSizesSource?: string | null;
};

/** In-memory host for unit tests (no Office/WPS process). */
export class MockHostAdapter implements HostAdapter {
  readonly kind = "office-js" as const;
  dynamicArrayFunctionsEnabled = false;
  workbookName = "Book1.xlsx";
  sheets: SheetInfo[] = [{ name: "Sheet1", index: 0, isActive: true }];
  cells = new Map<string, { values: CellValue[][]; formulas: string[][] }>();
  formats = new Map<string, RangeFormat>();
  tables: TableInfo[] = [];
  charts: ChartInfo[] = [];
  /** series keyed by sheet\\0chart; not part of ChartInfo public shape. */
  chartSeries = new Map<string, MockSeries[]>();
  chartAxes = new Map<string, ChartAxisInfo>();
  chartDataLabels = new Map<string, ChartDataLabelsInfo>();
  chartSeriesAxisGroups = new Map<string, ChartSeriesAxisGroupInfo>();
  shapes: import("../shared/host/shapeTypes").ShapeInfo[] = [];
  usedRangeAddress: string | null = "Sheet1!A1:B2";
  selection: SelectionInfo = {
    sheetName: "Sheet1",
    address: "Sheet1!A1",
    values: [[null]],
    formulas: [[""]],
  };
  failCapability: string | null = null;

  getRuntimeCapabilities() {
    return {
      dynamicArrayFunctionsEnabled: this.dynamicArrayFunctionsEnabled,
    };
  }

  async getStatus(): Promise<HostResult<HostStatus>> {
    if (this.failCapability === "host.status") {
      return unsupported("host.status", this.kind, "forced failure");
    }
    return ok({
      kind: this.kind,
      connected: true,
      hostName: "Mock Excel",
      workbookName: this.workbookName,
    });
  }

  async getSelection(): Promise<HostResult<SelectionInfo>> {
    return ok({ ...this.selection });
  }

  async readRange(
    sheetName: string,
    address: string,
    expand?: import("../shared/host/types").RangeExpandMode,
  ): Promise<HostResult<RangeData>> {
    const bare = address.includes("!") ? address.split("!")[1]! : address;
    const isSingle = !bare.includes(":") && !bare.includes(",");
    const effective =
      expand === undefined && isSingle ? ("spill" as const) : expand;
    if (effective && effective !== "none") {
      const hit = this.cells.get(key(sheetName, address));
      return ok({
        sheetName,
        address: `${sheetName}!${bare}:A3`,
        values: hit?.values ?? [["spill"]],
        formulas: hit?.formulas ?? [["=1"]],
        expanded: true,
        expandMode: effective,
      });
    }
    const hit = this.cells.get(key(sheetName, address));
    return ok({
      sheetName,
      address: `${sheetName}!${bare}`,
      values: hit?.values ?? [[null]],
      formulas: hit?.formulas ?? [[""]],
      expanded: false,
      expandMode: "none",
    });
  }

  async getFormulaContext(
    sheetName: string,
    address?: string,
  ): Promise<HostResult<import("../shared/host/types").FormulaContextData>> {
    const { absoluteA1FromOrigin } = await import("../shared/host/a1Address");
    const target = address?.trim() || "A1";
    const read = await this.readRange(sheetName, target, "none");
    if (!read.ok) return read;
    const origin = target.split(":")[0]!.replace(/^.*!/, "");
    const formulas = [];
    for (let r = 0; r < read.data.formulas.length; r += 1) {
      for (let c = 0; c < (read.data.formulas[r]?.length ?? 0); c += 1) {
        const formula = read.data.formulas[r][c] ?? "";
        if (!formula.startsWith("=")) continue;
        formulas.push({
          address: absoluteA1FromOrigin(origin, r, c),
          formula,
          value: read.data.values[r]?.[c] ?? null,
        });
      }
    }
    return ok({
      sheetName,
      address: origin.includes(":") ? origin : read.data.address.replace(/^.*!/, ""),
      formulas,
      cells: formulas,
    });
  }

  async copySheet(
    sheetName: string,
    newName?: string,
  ): Promise<HostResult<SheetInfo>> {
    const name = newName ?? `${sheetName}_Copy`;
    return this.addSheet(name);
  }

  async moveSheet(sheetName: string, position: number): Promise<HostResult<SheetInfo>> {
    const sheet = this.sheets.find((item) => item.name === sheetName);
    if (!sheet) return unsupported("sheet.move", this.kind, "missing sheet");
    // Public contract is 1-based.
    sheet.index = position;
    return ok({ ...sheet });
  }

  async writeRange(
    sheetName: string,
    address: string,
    values: CellValue[][],
  ): Promise<HostResult<RangeData>> {
    const formulas = values.map((row) => row.map(() => ""));
    this.cells.set(key(sheetName, address), { values, formulas });
    return this.readRange(sheetName, address);
  }

  async writeFormulas(
    sheetName: string,
    address: string,
    formulas: string[][],
  ): Promise<HostResult<RangeData>> {
    this.cells.set(key(sheetName, address), {
      values: formulas.map((row) => row.map(() => null)),
      formulas,
    });
    return this.readRange(sheetName, address);
  }

  async clearRange(
    sheetName: string,
    address: string,
  ): Promise<HostResult<{ cleared: string }>> {
    this.cells.delete(key(sheetName, address));
    return ok({ cleared: `${sheetName}!${address}` });
  }

  async insertRange(input: RangeInsertInput): Promise<HostResult<RangeMutationInfo>> {
    const address = normalizeSameSheetA1Range(
      input.sheetName,
      input.address,
      "range",
      "range operation",
    );
    return ok({
      sheetName: input.sheetName,
      address: `${input.sheetName}!${address}`,
      shift: input.shift,
      operation: "insert",
    });
  }

  async deleteRange(input: RangeDeleteInput): Promise<HostResult<RangeMutationInfo>> {
    const address = normalizeSameSheetA1Range(
      input.sheetName,
      input.address,
      "range",
      "range operation",
    );
    return ok({
      sheetName: input.sheetName,
      address: `${input.sheetName}!${address}`,
      shift: input.shift,
      operation: "delete",
    });
  }

  async autofitRange(input: RangeAutofitInput): Promise<HostResult<RangeAutofitInfo>> {
    const address = normalizeSameSheetA1Range(
      input.sheetName,
      input.address,
      "range",
      "range operation",
    );
    return ok({
      sheetName: input.sheetName,
      address: `${input.sheetName}!${address}`,
      direction: input.direction,
      columnWidth: 64,
      rowHeight: 18,
    });
  }

  async listSheets(): Promise<HostResult<SheetInfo[]>> {
    return ok([...this.sheets]);
  }

  async addSheet(sheetName: string): Promise<HostResult<SheetInfo>> {
    const sheet = { name: sheetName, index: this.sheets.length, isActive: false };
    this.sheets.push(sheet);
    return ok(sheet);
  }

  async renameSheet(sheetName: string, newName: string): Promise<HostResult<SheetInfo>> {
    const sheet = this.sheets.find((item) => item.name === sheetName);
    if (!sheet) return unsupported("sheet.rename", this.kind, "missing sheet");
    sheet.name = newName;
    return ok({ ...sheet });
  }

  async deleteSheet(sheetName: string): Promise<HostResult<{ deleted: string }>> {
    this.sheets = this.sheets.filter((item) => item.name !== sheetName);
    return ok({ deleted: sheetName });
  }

  async readFormat(sheetName: string, address: string): Promise<HostResult<RangeFormatData>> {
    return ok({
      sheetName,
      address: `${sheetName}!${address}`,
      format: { ...this.formats.get(key(sheetName, address)) },
    });
  }

  async writeFormat(
    sheetName: string,
    address: string,
    format: RangeFormat,
  ): Promise<HostResult<RangeFormatData>> {
    const prev = this.formats.get(key(sheetName, address)) ?? {};
    const next = { ...prev, ...format };
    this.formats.set(key(sheetName, address), next);
    return this.readFormat(sheetName, address);
  }

  async listTables(sheetName?: string): Promise<HostResult<TableInfo[]>> {
    return ok(
      this.tables.filter((table) => (sheetName ? table.sheetName === sheetName : true)),
    );
  }

  async createTable(input: {
    sheetName: string;
    address: string;
    name?: string;
    hasHeaders?: boolean;
  }): Promise<HostResult<TableInfo>> {
    const table: TableInfo = {
      name: input.name ?? `Table${this.tables.length + 1}`,
      sheetName: input.sheetName,
      address: `${input.sheetName}!${input.address}`,
      hasHeaders: input.hasHeaders !== false,
      showFilter: true,
    };
    this.tables.push(table);
    return ok(table);
  }

  async deleteTable(
    sheetName: string,
    tableName: string,
  ): Promise<HostResult<{ deleted: string }>> {
    this.tables = this.tables.filter(
      (table) => !(table.sheetName === sheetName && table.name === tableName),
    );
    return ok({ deleted: tableName });
  }

  async unlistTable(
    sheetName: string,
    tableName: string,
  ): Promise<HostResult<TableUnlistInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === sheetName && item.name === tableName,
    );
    if (!table) throw new Error(`table not found: ${tableName}`);
    this.tables = this.tables.filter(
      (item) => !(item.sheetName === sheetName && item.name === tableName),
    );
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      address: table.address,
      unlisted: true,
    });
  }


  private tableFilters = new Map<string, { enabled: boolean; columnIndex?: number; filterOn?: string }>();
  private tableSorts = new Map<string, { fields: Array<{ columnIndex: number; ascending: boolean }> }>();

  private tableKey(sheetName: string, tableName: string): string {
    return `${sheetName}::${tableName}`;
  }

  async getTableFilter(input: TableFilterGetInput): Promise<HostResult<TableFilterInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    const state = this.tableFilters.get(this.tableKey(input.sheetName, input.tableName));
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      enabled: state?.enabled === true,
      columnIndex: state?.columnIndex,
      filterOn: state?.filterOn as TableFilterInfo["filterOn"],
    });
  }

  async applyTableFilter(input: TableFilterApplyInput): Promise<HostResult<TableFilterInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    if (input.filterOn === "values" && (!input.values || input.values.length === 0)) {
      throw new Error("filterOn=values requires non-empty values[]");
    }
    if (input.filterOn === "custom" && (input.criterion1 == null || input.criterion1 === "")) {
      throw new Error("filterOn=custom requires criterion1");
    }
    if (
      (input.filterOn === "topItems" ||
        input.filterOn === "bottomItems" ||
        input.filterOn === "topPercent" ||
        input.filterOn === "bottomPercent") &&
      (input.threshold == null || input.threshold <= 0)
    ) {
      throw new Error(`${input.filterOn} requires positive threshold`);
    }
    this.tableFilters.set(this.tableKey(input.sheetName, input.tableName), {
      enabled: true,
      columnIndex: input.columnIndex,
      filterOn: input.filterOn,
    });
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      enabled: true,
      columnIndex: input.columnIndex,
      filterOn: input.filterOn,
    });
  }

  async clearTableFilter(input: TableFilterClearInput): Promise<HostResult<TableFilterInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    this.tableFilters.set(this.tableKey(input.sheetName, input.tableName), { enabled: false });
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      enabled: false,
    });
  }

  async getTableSort(input: TableSortGetInput): Promise<HostResult<TableSortInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    const state = this.tableSorts.get(this.tableKey(input.sheetName, input.tableName));
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      fields: state?.fields ?? [],
    });
  }

  async applyTableSort(input: TableSortApplyInput): Promise<HostResult<TableSortInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    const fields = input.fields.map((field) => ({
      columnIndex: field.columnIndex,
      ascending: field.ascending !== false,
    }));
    this.tableSorts.set(this.tableKey(input.sheetName, input.tableName), { fields });
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      fields,
    });
  }

  async clearTableSort(input: TableSortClearInput): Promise<HostResult<TableSortInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    this.tableSorts.set(this.tableKey(input.sheetName, input.tableName), { fields: [] });
    return ok({
      sheetName: table.sheetName,
      tableName: table.name,
      fields: [],
    });
  }


  /** sheetName -> { formulas[r][c], locked[r][c], protected } for formula protection tests */
  formulaProtectionSheets = new Map<
    string,
    {
      protected: boolean;
      formulas: string[][];
      locked: boolean[][];
      address: string;
    }
  >();

  private ensureFormulaSheet(sheetName: string) {
    let state = this.formulaProtectionSheets.get(sheetName);
    if (!state) {
      state = {
        protected: false,
        formulas: [[""]],
        locked: [[true]],
        address: `${sheetName}!A1`,
      };
      this.formulaProtectionSheets.set(sheetName, state);
    }
    return state;
  }

  async inspectFormulaProtection(
    input: FormulaProtectionInspectInput,
  ): Promise<HostResult<FormulaProtectionInspectInfo>> {
    const names =
      input.scope === "workbook"
        ? [...this.formulaProtectionSheets.keys()]
        : [input.sheetName ?? "Sheet1"];
    if (input.scope !== "workbook" && !input.sheetName) {
      throw new Error("sheetName is required for scope sheet|target");
    }
    if (input.scope === "target" && !input.range) {
      throw new Error("range is required for scope=target");
    }
    const sheets = [];
    let formulaCount = 0;
    let lockedFormulaCount = 0;
    for (const name of names.length ? names : ["Sheet1"]) {
      const state = this.ensureFormulaSheet(name);
      let fCount = 0;
      let lCount = 0;
      for (let r = 0; r < state.formulas.length; r++) {
        for (let c = 0; c < (state.formulas[r]?.length ?? 0); c++) {
          const f = state.formulas[r]![c] ?? "";
          if (typeof f === "string" && f.trim().startsWith("=")) {
            fCount += 1;
            if (state.locked[r]?.[c] === true) lCount += 1;
          }
        }
      }
      formulaCount += fCount;
      lockedFormulaCount += lCount;
      sheets.push({
        sheetName: name,
        address: input.range ? `${name}!${input.range}` : state.address,
        formulaCount: fCount,
        lockedFormulaCount: lCount,
        sheetProtected: state.protected,
        limitations: [],
      });
    }
    return ok({
      scope: input.scope,
      sheets,
      formulaCount,
      lockedFormulaCount,
      limitations: [],
    });
  }

  async manageFormulaProtection(
    input: FormulaProtectionManageInput,
  ): Promise<HostResult<FormulaProtectionManageInfo>> {
    // Intentionally never put password into return value.
    const unlockInputs = input.unlockInputs !== false;
    const protectSheet = input.command === "lock" ? input.protectSheet !== false : false;
    const names =
      input.scope === "workbook"
        ? [...this.formulaProtectionSheets.keys()]
        : [input.sheetName ?? "Sheet1"];
    if (input.scope !== "workbook" && !input.sheetName) {
      throw new Error("sheetName is required for scope sheet|target");
    }
    for (const name of names.length ? names : ["Sheet1"]) {
      const state = this.ensureFormulaSheet(name);
      if (state.protected) {
        // unprotect with optional password — do not store password
        state.protected = false;
      }
      if (input.command === "lock" && unlockInputs) {
        for (let r = 0; r < state.locked.length; r++) {
          for (let c = 0; c < (state.locked[r]?.length ?? 0); c++) {
            state.locked[r]![c] = false;
          }
        }
      }
      for (let r = 0; r < state.formulas.length; r++) {
        for (let c = 0; c < (state.formulas[r]?.length ?? 0); c++) {
          const f = state.formulas[r]![c] ?? "";
          if (typeof f === "string" && f.trim().startsWith("=")) {
            state.locked[r]![c] = input.command === "lock";
          }
        }
      }
      if (protectSheet) state.protected = true;
    }
    const protection = await this.inspectFormulaProtection({
      scope: input.scope,
      sheetName: input.sheetName,
      range: input.range,
    });
    if (!protection.ok) throw new Error("inspect failed after manage");
    return ok({
      command: input.command,
      scope: input.scope,
      unlockInputs: input.command === "lock" ? unlockInputs : false,
      protectSheet,
      protection: protection.data,
      verified: true,
      limitations:
        input.command === "lock" && unlockInputs
          ? ["unlockInputs: unlocked all cells in target range before locking formula cells (inputs outside range unchanged)"]
          : [],
    });
  }

  async updateTable(input: TableUpdateInput): Promise<HostResult<TableInfo>> {
    const table = this.tables.find(
      (item) => item.sheetName === input.sheetName && item.name === input.tableName,
    );
    if (!table) throw new Error(`table not found: ${input.tableName}`);
    if (input.newName != null) table.name = input.newName;
    if (input.style != null) table.style = input.style;
    if (input.showHeaders != null) table.hasHeaders = input.showHeaders;
    if (input.showTotals != null) table.showTotals = input.showTotals;
    if (input.showFilterButton != null) table.showFilter = input.showFilterButton;
    if (input.showBandedRows != null) table.showBandedRows = input.showBandedRows;
    if (input.showBandedColumns != null) table.showBandedColumns = input.showBandedColumns;
    if (input.showFirstColumn != null) table.showFirstColumn = input.showFirstColumn;
    if (input.showLastColumn != null) table.showLastColumn = input.showLastColumn;
    if (input.resizeAddress != null) {
      const bare = normalizeSameSheetA1Range(
        input.sheetName,
        input.resizeAddress,
        "resizeAddress",
        "table",
      );
      table.address = `${input.sheetName}!${bare}`;
    }
    return ok(table);
  }

  async listCharts(sheetName?: string): Promise<HostResult<ChartInfo[]>> {
    return ok(
      this.charts.filter((chart) => (sheetName ? chart.sheetName === sheetName : true)),
    );
  }

  async createChart(input: {
    sheetName: string;
    sourceRange: string;
    chartType?: ChartType;
    name?: string;
    title?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  }): Promise<HostResult<ChartInfo>> {
    const chart: ChartInfo = {
      name: input.name ?? `Chart${this.charts.length + 1}`,
      sheetName: input.sheetName,
      chartType: input.chartType ?? "column",
      title: input.title,
      left: input.left ?? 0,
      top: input.top ?? 0,
      width: input.width ?? 360,
      height: input.height ?? 240,
      style: 2,
      legendVisible: true,
    };
    this.charts.push(chart);
    this.chartSeries.set(chartKey(chart.sheetName, chart.name), [
      { name: "Series1", chartType: chart.chartType, smooth: false, bubbleSizesSource: null },
      { name: "Series2", chartType: chart.chartType, smooth: false, bubbleSizesSource: null },
    ]);
    return ok(chart);
  }

  async deleteChart(
    sheetName: string,
    chartName: string,
  ): Promise<HostResult<{ deleted: string }>> {
    this.charts = this.charts.filter(
      (chart) => !(chart.sheetName === sheetName && chart.name === chartName),
    );
    this.chartSeries.delete(chartKey(sheetName, chartName));
    return ok({ deleted: chartName });
  }

  async updateChart(input: {
    sheetName: string;
    chartName: string;
    newName?: string;
    chartType?: ChartType;
    title?: string;
    showTitle?: boolean;
    style?: number;
    showLegend?: boolean;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  }): Promise<HostResult<ChartInfo>> {
    const chart = this.charts.find(
      (item) => item.sheetName === input.sheetName && item.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    if (input.newName != null) {
      const prev = chartKey(input.sheetName, chart.name);
      const series = this.chartSeries.get(prev);
      chart.name = input.newName;
      if (series) {
        this.chartSeries.delete(prev);
        this.chartSeries.set(chartKey(input.sheetName, chart.name), series);
      }
    }
    if (input.chartType != null) chart.chartType = input.chartType;
    if (input.title != null) {
      chart.title = input.title;
      chart.titleVisible = input.showTitle !== false;
    } else if (input.showTitle != null) {
      chart.titleVisible = input.showTitle;
    }
    if (input.style != null) chart.style = input.style;
    if (input.showLegend != null) chart.legendVisible = input.showLegend;
    if (input.left != null) chart.left = input.left;
    if (input.top != null) chart.top = input.top;
    if (input.width != null) chart.width = input.width;
    if (input.height != null) chart.height = input.height;
    return ok(chart);
  }

  async listChartSeries(
    sheetName: string,
    chartName: string,
  ): Promise<HostResult<ChartSeriesInfo[]>> {
    const chart = this.charts.find((c) => c.sheetName === sheetName && c.name === chartName);
    if (!chart) throw new Error(`chart not found: ${chartName}`);
    const series = this.chartSeries.get(chartKey(sheetName, chartName)) ?? [];
    return ok(
      series.map((item, i) => ({
        index: i + 1,
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      })),
    );
  }

  async updateChartSeries(input: ChartSeriesUpdateInput): Promise<HostResult<ChartSeriesInfo>> {
    const series = this.chartSeries.get(chartKey(input.sheetName, input.chartName));
    if (!series) throw new Error(`chart not found: ${input.chartName}`);
    if (input.seriesIndex < 1 || input.seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${input.seriesIndex}`);
    }
    const item = series[input.seriesIndex - 1]!;
    if (input.newName != null) item.name = input.newName;
    if (input.chartType != null) item.chartType = input.chartType;
    if (input.smooth != null) item.smooth = input.smooth;
    return ok({
      index: input.seriesIndex,
      name: item.name,
      chartType: item.chartType,
      smooth: item.smooth,
    });
  }

  async deleteChartSeries(
    sheetName: string,
    chartName: string,
    seriesIndex: number,
  ): Promise<HostResult<ChartSeriesDeleteResult>> {
    const chart = this.charts.find((c) => c.sheetName === sheetName && c.name === chartName);
    if (!chart) throw new Error(`chart not found: ${chartName}`);
    const series = this.chartSeries.get(chartKey(sheetName, chartName));
    if (!series) throw new Error(`chart not found: ${chartName}`);
    if (seriesIndex < 1 || seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${seriesIndex}`);
    }
    series.splice(seriesIndex - 1, 1);
    return ok({
      sheetName,
      chartName: chart.name,
      deletedSeriesIndex: seriesIndex,
      remainingSeries: series.map((item, i) => ({
        index: i + 1,
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      })),
    });
  }

  async addChartSeries(input: ChartSeriesAddInput): Promise<HostResult<ChartSeriesAddResult>> {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const key = chartKey(input.sheetName, input.chartName);
    let series = this.chartSeries.get(key);
    if (!series) {
      series = [];
      this.chartSeries.set(key, series);
    }
    const item = {
      name: input.name ?? `Series${series.length + 1}`,
      chartType: chart.chartType,
      smooth: false,
      bubbleSizesSource: null as string | null,
    };
    series.push(item);
    return ok({
      sheetName: input.sheetName,
      chartName: chart.name,
      addedSeries: {
        index: series.length,
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      },
      dataBound: false,
    });
  }

  async updateChartSeriesValues(
    input: ChartSeriesValuesUpdateInput,
  ): Promise<HostResult<ChartSeriesValuesInfo>> {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const series = this.chartSeries.get(chartKey(input.sheetName, input.chartName));
    if (!series || input.seriesIndex < 1 || input.seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${input.seriesIndex}`);
    }
    const info: ChartSeriesValuesInfo = {
      sheetName: input.sheetName,
      chartName: chart.name,
      seriesIndex: input.seriesIndex,
      dataBound: true,
    };
    if (input.valuesRange != null) {
      const bare = normalizeSameSheetSourceRange(input.sheetName, input.valuesRange);
      info.valuesSource = `${input.sheetName}!${bare}`;
    }
    if (input.xValuesRange != null) {
      const bare = normalizeSameSheetSourceRange(input.sheetName, input.xValuesRange);
      info.xValuesSource = `${input.sheetName}!${bare}`;
    }
    return ok(info);
  }

  async updateChartSeriesBubbleSizes(
    input: ChartSeriesBubbleSizesUpdateInput,
  ): Promise<HostResult<ChartSeriesBubbleSizesInfo>> {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const series = this.chartSeries.get(chartKey(input.sheetName, input.chartName));
    if (!series || input.seriesIndex < 1 || input.seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${input.seriesIndex}`);
    }
    const bare = normalizeSameSheetSourceRange(input.sheetName, input.bubbleSizesRange);
    const source = `${input.sheetName}!${bare}`;
    const item = series[input.seriesIndex - 1]!;
    item.bubbleSizesSource = source;
    return ok({
      sheetName: input.sheetName,
      chartName: chart.name,
      seriesIndex: input.seriesIndex,
      bubbleSizesSource: item.bubbleSizesSource,
      dataBound: true,
    });
  }

  async getChartImage(input: ChartImageGetInput): Promise<HostResult<ChartImageInfo>> {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const dim =
      input.width != null || input.height != null
        ? `:${input.width ?? ""}x${input.height ?? ""}`
        : "";
    return ok({
      sheetName: chart.sheetName,
      chartName: chart.name,
      imageBase64: `bW9ja2ltYWdl${dim}`,
    });
  }

  async getRangeImage(input: RangeImageGetInput): Promise<HostResult<RangeImageInfo>> {
    const sheet = this.sheets.find((item) => item.name === input.sheetName);
    if (!sheet) throw new Error(`sheet not found: ${input.sheetName}`);
    const bare = input.range.includes("!") ? input.range.split("!")[1]! : input.range;
    return ok({
      sheetName: sheet.name,
      address: `${sheet.name}!${bare}`,
      imageBase64: "bW9ja3JhbmdlaW1hZ2U=",
    });
  }

  async updateChartSource(input: ChartSourceUpdateInput) {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const seriesBy = input.seriesBy ?? "auto";
    const sourceRange = normalizeSameSheetSourceRange(input.sheetName, input.sourceRange);
    const key = chartKey(input.sheetName, input.chartName);
    const bare = sourceRange;
    const parts = bare.includes(":") ? bare.split(":") : [bare, bare];
    const col = (a: string) => {
      const m = /^([A-Z]+)/i.exec(a);
      if (!m) return 1;
      let n = 0;
      for (const ch of m[1]!.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n;
    };
    const row = (a: string) => Number(/(\d+)$/.exec(a)?.[1] ?? "1");
    const cols = Math.max(1, col(parts[1]!) - col(parts[0]!) + 1);
    const rows = Math.max(1, row(parts[1]!) - row(parts[0]!) + 1);
    const count =
      seriesBy === "rows" ? Math.max(1, rows - 1) : Math.max(1, cols - 1);
    const prefix = seriesBy === "rows" ? "R" : "C";
    const next = Array.from({ length: count }, (_, i) => ({
      name: `${prefix}${i + 1}`,
      chartType: chart.chartType,
      smooth: false,
    }));
    this.chartSeries.set(key, next);
    return ok({
      sheetName: input.sheetName,
      chartName: chart.name,
      sourceRange,
      seriesBy,
      series: next.map((item, i) => ({
        index: i + 1,
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      })),
    });
  }

  private axisKey(
    sheetName: string,
    chartName: string,
    kind: string,
    group: string,
  ): string {
    return `${sheetName}\0${chartName}\0${kind}\0${group}`;
  }

  private ensureAxis(
    sheetName: string,
    chartName: string,
    kind: ChartAxisInfo["kind"],
    group: ChartAxisInfo["group"],
  ): ChartAxisInfo {
    const key = this.axisKey(sheetName, chartName, kind, group);
    let axis = this.chartAxes.get(key);
    if (!axis) {
      axis = {
        sheetName,
        chartName,
        kind,
        group,
        title: "",
        titleVisible: false,
        minimum: 0,
        maximum: 100,
        majorUnit: 10,
        numberFormat: "General",
        reverse: false,
      };
      this.chartAxes.set(key, axis);
    }
    return axis;
  }

  async updateChartAxis(input: ChartAxisUpdateInput) {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const group = input.group ?? "primary";
    const axis = this.ensureAxis(input.sheetName, input.chartName, input.kind, group);
    axis.sheetName = input.sheetName;
    axis.chartName = chart.name;
    if (input.title !== undefined) {
      axis.title = input.title;
      axis.titleVisible = input.title !== "";
    }
    if (input.minimum !== undefined) axis.minimum = input.minimum;
    if (input.maximum !== undefined) axis.maximum = input.maximum;
    if (input.majorUnit !== undefined) axis.majorUnit = input.majorUnit;
    if (input.numberFormat !== undefined) axis.numberFormat = input.numberFormat;
    if (input.reverse !== undefined) axis.reverse = input.reverse;
    return ok({ ...axis });
  }

  async updateChartDataLabels(input: ChartDataLabelsUpdateInput) {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const series = this.chartSeries.get(chartKey(input.sheetName, input.chartName));
    if (!series || input.seriesIndex < 1 || input.seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${input.seriesIndex}`);
    }
    const key = `${input.sheetName}\0${input.chartName}\0${input.seriesIndex}`;
    let labels = this.chartDataLabels.get(key);
    if (!labels) {
      labels = {
        sheetName: input.sheetName,
        chartName: chart.name,
        seriesIndex: input.seriesIndex,
        enabled: false,
        showValue: false,
        showCategoryName: false,
        showSeriesName: false,
        numberFormat: "General",
      };
      this.chartDataLabels.set(key, labels);
    }
    labels.chartName = chart.name;
    if (input.enabled !== undefined) labels.enabled = input.enabled;
    if (input.showValue !== undefined) labels.showValue = input.showValue;
    if (input.showCategoryName !== undefined) labels.showCategoryName = input.showCategoryName;
    if (input.showSeriesName !== undefined) labels.showSeriesName = input.showSeriesName;
    if (input.numberFormat !== undefined) labels.numberFormat = input.numberFormat;
    return ok({ ...labels });
  }

  async updateChartSeriesAxisGroup(input: ChartSeriesAxisGroupUpdateInput) {
    const chart = this.charts.find(
      (c) => c.sheetName === input.sheetName && c.name === input.chartName,
    );
    if (!chart) throw new Error(`chart not found: ${input.chartName}`);
    const series = this.chartSeries.get(chartKey(input.sheetName, input.chartName));
    if (!series || input.seriesIndex < 1 || input.seriesIndex > series.length) {
      throw new Error(`seriesIndex out of range: ${input.seriesIndex}`);
    }
    const key = `${input.sheetName}\0${input.chartName}\0${input.seriesIndex}`;
    const info: ChartSeriesAxisGroupInfo = {
      sheetName: input.sheetName,
      chartName: chart.name,
      seriesIndex: input.seriesIndex,
      axisGroup: input.axisGroup,
    };
    this.chartSeriesAxisGroups.set(key, info);
    return ok({ ...info });
  }

  async inspectWorkbook(): Promise<HostResult<WorkbookInspectInfo>> {
    const active = this.sheets.find((sheet) => sheet.isActive) ?? this.sheets[0];
    return ok({
      workbookName: this.workbookName,
      activeSheetName: active?.name ?? "",
      sheetCount: this.sheets.length,
      usedRangeAddress: this.usedRangeAddress,
      sheets: [...this.sheets],
    });
  }

  private cfRules = new Map<
    string,
    import("../shared/host/types").ConditionalFormatInfo[]
  >();
  private dvRules = new Map<string, import("../shared/host/types").DataValidationRule | null>();

  async listConditionalFormats(sheetName: string, range: string) {
    return ok(this.cfRules.get(key(sheetName, range)) ?? []);
  }

  async addConditionalFormat(input: {
    sheetName: string;
    range: string;
    rule: import("../shared/host/types").ConditionalFormatRule;
  }) {
    const id = `cf_${Date.now()}`;
    const info: import("../shared/host/types").ConditionalFormatInfo = {
      id,
      sheetName: input.sheetName,
      range: input.range,
      kind: input.rule.kind,
      summary: `${input.rule.kind}:${id}`,
    };
    const k = key(input.sheetName, input.range);
    const list = this.cfRules.get(k) ?? [];
    list.push(info);
    this.cfRules.set(k, list);
    return ok(info);
  }

  async deleteConditionalFormat(sheetName: string, range: string, id: string) {
    const k = key(sheetName, range);
    const list = (this.cfRules.get(k) ?? []).filter((item) => item.id !== id);
    this.cfRules.set(k, list);
    return ok({ deleted: id });
  }

  async readDataValidation(sheetName: string, range: string) {
    return ok({
      sheetName,
      range,
      rule: this.dvRules.get(key(sheetName, range)) ?? null,
    });
  }

  async writeDataValidation(input: {
    sheetName: string;
    range: string;
    rule: import("../shared/host/types").DataValidationRule;
  }) {
    this.dvRules.set(key(input.sheetName, input.range), input.rule);
    return ok({
      sheetName: input.sheetName,
      range: input.range,
      rule: input.rule,
    });
  }

  async clearDataValidation(sheetName: string, range: string) {
    this.dvRules.set(key(sheetName, range), null);
    return ok({ cleared: `${sheetName}!${range}` });
  }

  private structure = createMockStructureState();
  getSheetVisibility = this.structure.getSheetVisibility.bind(this.structure);
  setSheetVisibility = this.structure.setSheetVisibility.bind(this.structure);
  getSheetProtection = this.structure.getSheetProtection.bind(this.structure);
  protectSheet = this.structure.protectSheet.bind(this.structure);
  unprotectSheet = this.structure.unprotectSheet.bind(this.structure);
  listNamedRanges = this.structure.listNamedRanges.bind(this.structure);
  createNamedRange = this.structure.createNamedRange.bind(this.structure);
  updateNamedRange = this.structure.updateNamedRange.bind(this.structure);
  deleteNamedRange = this.structure.deleteNamedRange.bind(this.structure);
  getSheetDisplay = this.structure.getSheetDisplay.bind(this.structure);
  setSheetDisplay = this.structure.setSheetDisplay.bind(this.structure);
  getSheetFreeze = this.structure.getSheetFreeze.bind(this.structure);
  setSheetFreeze = this.structure.setSheetFreeze.bind(this.structure);
  getSheetPageLayout = this.structure.getSheetPageLayout.bind(this.structure);
  setSheetPageLayout = this.structure.setSheetPageLayout.bind(this.structure);

  async listShapes(sheetName?: string) {
    return ok(
      this.shapes.filter((shape) => (sheetName ? shape.sheetName === sheetName : true)),
    );
  }

  async createShape(input: import("../shared/host/shapeTypes").ShapeCreateInput) {
    const shape: import("../shared/host/shapeTypes").ShapeInfo = {
      name: input.name ?? `Shape${this.shapes.length + 1}`,
      sheetName: input.sheetName,
      type: input.kind === "geometric" ? "GeometricShape" : "GeometricShape",
      geometricShapeType: input.kind === "geometric" ? input.geometricType : null,
      left: input.left ?? 0,
      top: input.top ?? 0,
      width: input.width ?? 100,
      height: input.height ?? 100,
      visible: true,
      // Match Office.js: text only when non-empty hasText; else null.
      text:
        input.kind === "textBox" && input.text != null && input.text.length > 0
          ? input.text
          : null,
    };
    this.shapes.push(shape);
    return ok(shape);
  }

  async deleteShape(sheetName: string, shapeName: string) {
    this.shapes = this.shapes.filter(
      (shape) => !(shape.sheetName === sheetName && shape.name === shapeName),
    );
    return ok({ deleted: shapeName });
  }

  async updateShape(input: import("../shared/host/shapeTypes").ShapeUpdateInput) {
    const shape = this.shapes.find(
      (item) => item.sheetName === input.sheetName && item.name === input.shapeName,
    );
    if (!shape) throw new Error(`shape not found: ${input.shapeName}`);
    if (input.newName != null) shape.name = input.newName;
    if (input.left != null) shape.left = input.left;
    if (input.top != null) shape.top = input.top;
    if (input.width != null) shape.width = input.width;
    if (input.height != null) shape.height = input.height;
    if (input.visible != null) shape.visible = input.visible;
    if (input.text != null) {
      shape.text = input.text.length > 0 ? input.text : null;
    }
    return ok(shape);
  }
}
