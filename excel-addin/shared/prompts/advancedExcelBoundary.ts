import {
  resolveOfficeAdvancedIntents,
  type PromptRoutingContext,
} from "./promptRouting";

/**
 * Hard add-in runtime boundary. Always appended after synced desktop scenarios
 * so macro/base/office-tools narrative cannot imply unavailable tools are live.
 */
export function buildAddInHardBoundary(): string {
  return [
    "## 本加载项运行时能力边界（覆盖同步桌面提示词中的冲突叙述）",
    "- 宏：`macro.detect` / `macro.write` / `macro.run` 以及 VBA/JSA 写入与运行在本加载项均为 **unsupported**（Office.js 与 WPS JSA 宿主均无实现）；不得调用或伪造成功。",
    "- 文件级 Open XML：不读写 `.xlsx` 包体、不提供 C# Open XML 路径；相关能力 **unsupported**。",
    "- 运行时隔离：COM / .NET Worker / Electron IPC / child_process **禁止且 unsupported**。",
    "- UI：任务窗格外控件、UserForm、ActiveX、自定义功能区/菜单 **unsupported**；不得建议 `ui.addControl` / `ui.listControls` 作为本加载项能力。",
    "- Power Query / 透视表 / 切片器：**unsupported**，不得伪造成功。",
    "- 仅已注册工具可执行；同步自桌面的宏/Open XML/COM 叙述若与本段冲突，以本段为准。",
  ].join("\n");
}

/**
 * Add-in Excel tool boundary for office-tools scenario.
 * Values / formulas / format / basic table & chart are first-class; PQ/Pivot stay unsupported.
 */
export function buildAdvancedExcelBoundary(context: PromptRoutingContext = {}): string {
  const intents = resolveOfficeAdvancedIntents(context);
  const rules = [
    "Excel 工具边界：值与固定汇总用 `range.write`；公式用 `formula.write`；读取含公式单元格用 `formula.context`（range 可省略→UsedRange）；区域读取用 `range.read`（可选 expand:none|spill|currentArray|currentRegion；省略 expand 且单单元格时自动 spill，显式 none 不扩展）；字体/填充/数字格式/对齐/换行用 `range.format.write`；条件格式用 `conditionalFormat.list/add/delete`（cellValue/custom）；数据验证用 `dataValidation.read/write/clear`（list/wholeNumber）；工作表结构用 `sheet.operation`（add/rename/delete/copy/move，position 为 1-based）；可见性用 `sheet.visibility.get/set`（visible|hidden|veryHidden）；显示属性用 `sheet.display.get/set`（tabColor 空串=自动/#RRGGBB、showGridlines、showHeadings）；冻结窗格用 `sheet.freeze.get/set`（rows|columns|at|clear）；页面布局/打印设置用 `sheet.pageLayout.get/set`（ExcelApi 1.9：orientation/margins.{top,bottom,left,right,header,footer} points/printArea/paperSize a3|a4|a5|letter|legal/fitToPagesWide|Tall/zoomScale/draft/pageOrder downThenOver|overThenDown/firstPageNumber≥1 等；fit 与 zoomScale 互斥；headers/footers 文本内容/page breaks/clear/fitToOnePage → unsupported）；保护用 `sheet.protection.get/protect/unprotect`（password 仅请求内存）；命名区域用 `namedRange.list/create/update/delete`（scope workbook|worksheet）；基础表格用 `table.list/create/delete/update/unlist`（update 仅浅层 name/style/headers/totals/filter；unlist=ExcelApi 1.2 convertToRange 保留数据转普通区域；delete 仍硬删除；resize/filter/sort/banded/highlight → unsupported）；基础图表用 `chart.list/create/delete/update`（type: column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers；create enum ExcelApi 1.1；chart/series chartType 属性 ExcelApi 1.7；update 浅层 name/type/title/style/showLegend/pos/size）；图表 series 用 `chart.series.list/update`（仅 name/chartType/smooth；seriesIndex 1-based；**不处理 dataLabels**）；图表数据源用 `chart.source.update`（同表 A1 sourceRange + seriesBy auto|rows|columns，默认 auto）；图表坐标轴用 `chart.axes.update`（kind category|value；group primary|secondary 默认 primary；title/min/max/majorUnit/numberFormat/reverse）；图表数据标签用 `chart.series.dataLabels.update`（seriesIndex 1-based；enabled-only ExcelApi 1.7 hasDataLabels（不访问 dataLabels）; showValue/showCategoryName/showSeriesName/numberFormat ExcelApi 1.8 完整快照；enabled=false 不可与其它标签字段同传；showPercentage/showBubbleSize/delete/position/format/leaderLines 等 → unsupported）；图表 series 轴组用 `chart.series.axisGroup.update`（seriesIndex 1-based；axisGroup primary|secondary）；图表 series 删除用 `chart.series.delete`（seriesIndex 1-based；回读 remainingSeries）；图表空 series 创建用 `chart.series.add`（可选 name；dataBound=false，设置 values/xValues 前不可见）；图表 series 数据源绑定用 `chart.series.values.update`（valuesRange/xValuesRange 同表 A1；ExcelApi 1.15 getDimensionDataSourceString 真源回读；dataBound=true）；图表 series bubble sizes 用 `chart.series.bubbleSizes.update`（bubbleSizesRange 同表 A1；ExcelApi 1.15 getDimensionDataSourceString('BubbleSizes') 真源回读；仅 bubble chart；dataBound=true）；图表图像用 `chart.image.get`（ExcelApi 1.2 Chart.getImage 内存 Base64；可选 width/height；无路径/PDF）；形状用 `shape.list/create/delete/update`（create: geometric 白名单 rectangle|ellipse|triangle|diamond|rightArrow 或 textBox；update 浅层 newName/left/top/width/height/text/visible）。series formula/categoryFormula/format/trendlines/bubbleScale/showBubbleSize、跨表 source、seriesAxis/displayUnit/log/gridlines、export、复杂布局、image/line/group/fill/rotation/zOrder/scale/copyTo/getAsImage/preview → unsupported。数据量不是升级到高级对象的理由。",
    "宿主差异：Microsoft Excel（Office.js）支持 expand/format/conditionalFormat/dataValidation/table/table.unlist/chart/chart.series/chart.source/chart.axes/chart.series.dataLabels/chart.series.axisGroup/chart.series.delete/chart.series.add/chart.series.values.update/chart.series.bubbleSizes.update/chart.image.get/shape/copy/move/visibility/protection/namedRange/table.update/chart.update/sheet.display/sheet.freeze/sheet.pageLayout；WPS JSA 对本批 expand、format、conditionalFormat、dataValidation、table、table.unlist、chart、chart.series、chart.source、chart.axes、chart.series.dataLabels、chart.series.axisGroup、chart.series.delete、chart.series.add、chart.series.values.update、chart.series.bubbleSizes.update、chart.image.get、shape、copy/move、visibility、protection、namedRange、table.update、chart.update、sheet.display、sheet.freeze、sheet.pageLayout 返回 unsupported，不得伪造成功。",
    "硬性 unsupported：宏写入/运行、文件级 Open XML、COM/.NET/Electron、UI 控件/UserForm/菜单、Power Query/Pivot（与宿主无关，本加载项均不实现）。",
  ];
  if (intents.has("refreshable-etl")) {
    rules.push(
      "本轮检测到外部/多来源可刷新 ETL 意图，但本加载项批次尚未实现 Power Query 工具，相关能力为 unsupported，不得伪造成功。",
    );
  }
  if (intents.has("interactive-pivot")) {
    rules.push(
      "本轮检测到交互式透视意图，但本加载项批次尚未实现透视表/切片器工具，相关能力为 unsupported，不得伪造成功。",
    );
  }
  if (!intents.has("refreshable-etl") && !intents.has("interactive-pivot")) {
    rules.push("本轮未检测到 Power Query / 交互透视高级意图，相关 operation 不向模型开放。");
  }
  return `- ${rules.join(" ")}`;
}
