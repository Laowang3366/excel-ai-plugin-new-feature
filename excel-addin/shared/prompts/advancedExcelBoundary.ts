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
    "- 作用域：仅当前活动 Excel/WPS 工作簿；可用 `workbook.save` 原地保存当前簿；不得声称可打开/创建/saveAs/切换任意磁盘路径上的其他工作簿。",
    "- 宏：`macro.detect` / `macro.write` / `macro.run` 以及 VBA/JSA 写入与运行均为 **unsupported**；不得调用或伪造成功。",
    "- 文件级 Open XML / C# 文件服务：不读写 `.xlsx` 包体；相关能力 **unsupported**。",
    "- 运行时隔离：COM / .NET Worker / Electron IPC / child_process **禁止且 unsupported**。",
    "- UI：任务窗格外控件、UserForm、ActiveX、自定义功能区/菜单 **unsupported**；不得建议 `ui.addControl` / `ui.listControls`。",
    "- Power Query / 切片器：**unsupported**，不得伪造成功。Office.js 可用 `pivot.list`/`pivot.create`（ExcelApi 1.8）与 `pivot.refresh`（ExcelApi 1.3，当前工作簿；需 advancedIntent=interactive-pivot）；`refreshConnections` 桌面参数加载项不支持；WPS 透视表 **typed unsupported**；不得伪造成功。",
    "- 事务备份 / `office.workflow` / `office.transaction` / 跨 Word·PPT·PDF 导出：**unsupported**。",
    "- 仅本轮已注册且可见的工具可执行；与桌面端冲突的叙述以本段与适配模板为准。",
  ].join("\n");
}

/**
 * Add-in Excel tool boundary for office-tools scenario.
 * Values / formulas / format / basic table & chart are first-class; PQ/Pivot stay unsupported.
 */
export function buildAdvancedExcelBoundary(context: PromptRoutingContext = {}): string {
  const intents = resolveOfficeAdvancedIntents(context);
  const rules = [
    "Excel 工具边界：值与固定汇总用 `range.write`；公式用 `formula.write`；读取含公式单元格用 `formula.context`；公式保护用 `formula.protection.inspect`（safe）/`formula.protection.manage`（dangerous，仅锁公式单元格；unlockInputs 仅目标范围；password 不回传；WPS unsupported）；公式治理用 `formula.dependencies.inspect`（safe，文本解析依赖图+limitations）、`formula.references.repair`/`formula.convertToValues`/`formula.backups.restore`（dangerous，隐藏表 WENGGE_FORMULA_BACKUP_V1）、`formula.backups.inspect`（safe）；repair 计划不全则不写；convert 强制备份（range 可省略→UsedRange）；区域读取用 `range.read`（可选 expand:none|spill|currentArray|currentRegion；省略 expand 且单单元格时自动 spill，显式 none 不扩展）；字体/填充/数字格式/对齐/换行用 `range.format.write`；行列结构变更用 `range.insert`（shift down|right，ExcelApi 1.1）、`range.delete`（shift up|left，ExcelApi 1.1）和 `range.autofit`（rows|columns|both，ExcelApi 1.2，回读实际行高/列宽）；条件格式用 `conditionalFormat.list/add/delete`（ExcelApi 1.6；list 保留真实 hostType，add 仅 cellValue/custom 完整运算符）；数据验证用 `dataValidation.read/write/clear`（ExcelApi 1.8；list/wholeNumber/decimal/date/time/textLength/custom；list 内联与区域源互斥；Inconsistent/MixedCriteria 不构造假 rule；errorAlert/prompt 不实现）；工作表结构用 `sheet.operation`（add/rename/delete/copy/move，position 为 1-based）；可见性用 `sheet.visibility.get/set`（visible|hidden|veryHidden）；显示属性用 `sheet.display.get/set`（tabColor 空串=自动/#RRGGBB、showGridlines、showHeadings）；冻结窗格用 `sheet.freeze.get/set`（rows|columns|at|clear）；页面布局/打印设置用 `sheet.pageLayout.get/set`（ExcelApi 1.9：orientation/margins.{top,bottom,left,right,header,footer} points/headers|footers default 页 left|center|right 文本（空串清除）/manual horizontalPageBreaks|verticalPageBreaks bare A1（append，[] no-op）+ clearPageBreaks/printArea/printTitleRows|Columns/paperSize a3|a4|a5|letter|legal/fitToPagesWide|Tall/fitToOnePageWide|Tall/zoomScale/draft/pageOrder downThenOver|overThenDown/firstPageNumber≥1 等；fit/fitToOnePage 与 zoomScale 互斥；printTitleRows|Columns 别名 repeatRows|Columns；fitToOnePageWide|Tall→fit pages 1；自动分页、first/even/odd pages、headersFooters state/图片、printArea/titles clear → unsupported）；保护用 `sheet.protection.get/protect/unprotect`（password 仅请求内存）；命名区域用 `namedRange.list/create/update/delete`（scope workbook|worksheet）；基础表格用 `table.list/create/delete/update/unlist`（update 支持 name/style/headers/totals/filterButton/showBandedRows/showBandedColumns/showFirstColumn/showLastColumn/resizeAddress；resizeAddress 为同表单区域 A1，ExcelApi 1.13；banded/首末列 ExcelApi 1.3；unlist=ExcelApi 1.2 convertToRange；delete 仍硬删除）；表格筛选用 `table.filter.get/apply/clear`（AutoFilter apply/clear ExcelApi 1.2，enabled 1.9；columnIndex 1-based；filterOn values|custom|top/bottom items/percent；cellColor/fontColor/icon/dynamic → unsupported）；表格排序用 `table.sort.get/apply/clear`（ExcelApi 1.2；fields≤3；columnIndex 1-based；仅 value 排序；颜色/图标排序 → unsupported）；基础图表用 `chart.list/create/delete/update`（type: column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers；create enum ExcelApi 1.1；chart/series chartType 属性 ExcelApi 1.7；update 浅层 name/type/title/style/showLegend/pos/size）；图表 series 用 `chart.series.list/update`（仅 name/chartType/smooth；seriesIndex 1-based；**不处理 dataLabels**）；图表数据源用 `chart.source.update`（同表或跨表 sourceRange：A1 / Sheet2!A1 / `'Sheet 2'!A1`；seriesBy auto|rows|columns，默认 auto；拒绝外部工作簿/3D/多区域/结构化引用）；图表坐标轴用 `chart.axes.update`（kind category|value；group primary|secondary 默认 primary；title/min/max/majorUnit/numberFormat/reverse；displayUnit/customDisplayUnit/scaleType/logBase/showDisplayUnitLabel ExcelApi 1.7；majorGridlinesVisible/minorGridlinesVisible ExcelApi 1.1；写后宿主回读）；图表数据标签用 `chart.series.dataLabels.update`（seriesIndex 1-based；enabled-only ExcelApi 1.7 hasDataLabels（不访问 dataLabels）; showValue/showCategoryName/showSeriesName/numberFormat ExcelApi 1.8 完整快照；enabled=false 不可与其它标签字段同传；showPercentage/showBubbleSize/delete/position/format/leaderLines 等 → unsupported）；图表 series 轴组用 `chart.series.axisGroup.update`（seriesIndex 1-based；axisGroup primary|secondary）；图表 series 删除用 `chart.series.delete`（seriesIndex 1-based；回读 remainingSeries）；图表空 series 创建用 `chart.series.add`（可选 name；dataBound=false，设置 values/xValues 前不可见）；图表 series 数据源绑定用 `chart.series.values.update`（valuesRange/xValuesRange 同表 A1；ExcelApi 1.15 getDimensionDataSourceString 真源回读；dataBound=true）；图表 series bubble sizes 用 `chart.series.bubbleSizes.update`（bubbleSizesRange 同表 A1；ExcelApi 1.15 getDimensionDataSourceString('BubbleSizes') 真源回读；仅 bubble chart；dataBound=true）；图表 series 趋势线用 `chart.series.trendlines.list/add/update/delete`（seriesIndex+trendlineIndex 1-based；type linear|exponential|logarithmic|movingAverage|polynomial|power；ExcelApi 1.7 核心字段 + 1.8 period/equation；intercept 可写空串表示自动；写后宿主回读；WPS unsupported）；图表 series 标记用 `chart.series.markers.update`（markerStyle/markerSize 2-72/markerBackgroundColor|markerForegroundColor #RRGGBB；ExcelApi 1.7 写后回读；WPS unsupported）；图表图像用 `chart.image.get`（ExcelApi 1.2 Chart.getImage 内存 Base64；可选 width/height；无路径/PDF）；区域 PNG 用 `range.image.get`（ExcelApi **1.7** `Range.getImage` 内存 Base64；无路径/PDF/MIME/width/height；WPS unsupported）；形状用 `shape.list/create/delete/update`（create: geometric 白名单 rectangle|ellipse|triangle|diamond|rightArrow 或 textBox；update 浅层 newName/left/top/width/height/text/visible）。series formula/categoryFormula/per-point markers/trendline-format/bubbleScale/showBubbleSize、seriesAxis/axis-format、export、复杂布局、image/line/group/fill/rotation/zOrder/scale/copyTo/getAsImage/preview → unsupported。数据量不是升级到高级对象的理由。",
    "宿主差异：Microsoft Excel（Office.js）支持 `workbook.save`、expand none|spill|currentArray|currentRegion、format、range.insert/range.delete/range.autofit、conditionalFormat、dataValidation、table/table.unlist/table.filter/table.sort、formula.protection、formula.dependencies/formula.references/formula.convertToValues/formula.backups、chart 全系（含跨表 chart.source）、range.image、shape、copy/move、visibility、protection、namedRange、table.update、chart.update、sheet.display/sheet.freeze/sheet.pageLayout、`pivot.list`/`pivot.create`（ExcelApi 1.8）/`pivot.refresh`（ExcelApi 1.3）。WPS JSA（代码已落地 + 成员探测 + mock 单测；**非**官方 JSA 合同、**非**真机侧载）：**可调用** `workbook.save`（ActiveWorkbook.Save 成员探测）、`range.read` expand **currentRegion**（及 none）、`range.format.read/write`、`range.autofit`、`range.insert`/`range.delete`、`sheet.operation` copy/move、`sheet.visibility.get/set`、`sheet.protection.get/protect/unprotect`（password 仅请求内存、不得回显）、`namedRange.list/create/update/delete`、公式治理 `formula.dependencies.inspect`/`formula.references.repair`/`formula.convertToValues`/`formula.backups.inspect|restore`、`conditionalFormat.list/add/delete`、`dataValidation.read/write/clear`——成员缺失时必须尊重宿主 typed unsupported，不得改走 COM/.NET/Shell 或伪造成功。WPS **仍 typed unsupported**（勿扩大声明）：expand **spill|currentArray**、`formula.protection.*`、table/filter/sort/unlist、chart 全系、range.image、shape、sheet.display/sheet.freeze/sheet.pageLayout、`pivot.*`。",
    "透视表用 `pivot.list`/`pivot.create`（Office.js ExcelApi 1.8）/`pivot.refresh`（ExcelApi 1.3；create/refresh 需 advancedIntent=interactive-pivot；≥1 字段；空 destination→Pivots lastBottom+3；`refreshConnections` 非桌面对等、拒绝 true；切片器 unsupported；WPS unsupported）。对象清单优先 `workbook.objects.inspect`（safe；按分类 available|unsupported|failed；maxItemsPerCategory 截断且保留 totalCount），避免连续 list 膨胀上下文。",
    "硬性 unsupported：宏写入/运行、文件级 Open XML、COM/.NET/Electron、UI 控件/UserForm/菜单、Power Query、切片器、WPS 透视表；Office.js 透视表仅 `pivot.list`/`pivot.create`/`pivot.refresh`（无 slicer/PQ）。",
  ];
  if (intents.has("refreshable-etl")) {
    rules.push(
      "本轮检测到外部/多来源可刷新 ETL 意图，但本加载项批次尚未实现 Power Query 工具，相关能力为 unsupported，不得伪造成功。",
    );
  }
  if (intents.has("interactive-pivot")) {
    rules.push(
      "本轮检测到交互式透视意图：Office.js 使用 `pivot.list`（safe）/`pivot.create`/`pivot.refresh`（需 advancedIntent=interactive-pivot；写后回读；destination 空→Pivots 表；refreshConnections=true 拒绝（非桌面对等）；切片器仍 unsupported）；WPS 透视表 typed unsupported，不得伪造成功。普通汇总仍用 range/table，勿无故升级为透视表。",
    );
  }
  if (!intents.has("refreshable-etl") && !intents.has("interactive-pivot")) {
    rules.push("本轮未检测到 Power Query / 交互透视高级意图，相关 operation 不向模型开放。");
  }
  return `- ${rules.join(" ")}`;
}
