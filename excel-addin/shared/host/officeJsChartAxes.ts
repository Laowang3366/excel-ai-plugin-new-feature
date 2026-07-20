import { withExcel } from "./officeJsRuntime";
import type { ChartAxisGroup, ChartAxisInfo, ChartAxisUpdateInput } from "./chartAxisTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";
import {
  applyAxisWrites,
  EVIDENCE_17,
  EVIDENCE_18,
  EVIDENCE_19,
  GROUP_OFFICE,
  isSetSupported,
  KIND_OFFICE,
  needsExcelApi17,
  needsExcelApi18,
  needsExcelApi19,
  requireLoadedString,
  toAxisInfo,
  type ExcelChartAxes,
} from "./officeJsChartAxesMap";

/** Update chart axis fields; write → sync → load+sync real snapshot. */
export async function officeJsUpdateChartAxis(
  input: ChartAxisUpdateInput,
): Promise<HostResult<ChartAxisInfo>> {
  if (needsExcelApi17(input) && !isSetSupported("1.7")) {
    return unsupported(
      "chart.axes.update",
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      EVIDENCE_17,
    );
  }
  if (needsExcelApi18(input) && !isSetSupported("1.8")) {
    return unsupported(
      "chart.axes.update",
      "office-js",
      "ExcelApi 1.8 is not supported for ChartAxis.position/setPositionAt",
      EVIDENCE_18,
    );
  }
  if (needsExcelApi19(input) && !isSetSupported("1.9")) {
    return unsupported(
      "chart.axes.update",
      "office-js",
      "ExcelApi 1.9 is not supported for ChartAxis.linkNumberFormat",
      EVIDENCE_19,
    );
  }

  return withExcel("chart.axes.update", async (context) => {
    const group: ChartAxisGroup = input.group ?? "primary";
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      axes: ExcelChartAxes;
      load(props: string): void;
    };
    const axis = chart.axes.getItem(KIND_OFFICE[input.kind], GROUP_OFFICE[group]);
    applyAxisWrites(axis, input);
    await context.sync();
    const flags = {
      v17: isSetSupported("1.7"),
      v18: isSetSupported("1.8"),
      v19: isSetSupported("1.9"),
    };
    chart.load("name");
    let props =
      "type,axisGroup,minimum,maximum,majorUnit,minorUnit,numberFormat,reversePlotOrder";
    if (flags.v17) {
      props +=
        ",displayUnit,customDisplayUnit,scaleType,logBase,showDisplayUnitLabel,majorTickMark,minorTickMark,tickLabelPosition";
    }
    if (flags.v18) props += ",position,positionAt";
    if (flags.v19) props += ",linkNumberFormat";
    axis.load(props);
    axis.title.load("text,visible");
    axis.majorGridlines.load("visible");
    axis.minorGridlines.load("visible");
    await context.sync();
    return toAxisInfo(
      axis,
      input.sheetName,
      requireLoadedString(chart.name, "Chart.name"),
      flags,
    );
  });
}
