# Excel capability matrix (add-in validation)

Evidence columns:

- **office-js**: implemented against Office.js `Excel.run` / RequestContext in this package
- **wps-jsa**: only members with in-repo or strictly checked contracts; otherwise typed unsupported
- **desktop-source**: current Electron + .NET Worker tool surface (not used at runtime here)

| Capability family | Capability | office-js | wps-jsa | desktop-source | Evidence / notes |
|---|---|---|---|---|---|
| host | connection status | implemented | implemented | `office.connection.status` | `shared/host/*` |
| selection | selection.get | implemented | implemented* | `selection.get` | *Assumes `Application.Selection` |
| range | range.read/write/clear | implemented | implemented* | `range.*` | *Assumes Value2/Formula/Clear |
| range | range.read expand spill/currentArray/currentRegion | implemented | **unsupported** | expand modes | Office.js APIs; omit expand on single cell → spill (desktop parity) |
| formula | formula.read / formula.write | implemented | implemented* | formula tools | writeFormulas primitive |
| formula | formula.context | implemented | implemented* | `formula.context` | Desktop shape `{sheetName,address,formulas[]}`; range optional → UsedRange |
| sheet | list/add/rename/delete | implemented | implemented* | `sheet.operation` | |
| sheet | sheet.operation copy | implemented | **unsupported** | `sheet.operation` copy | Office.js `Worksheet.copy` |
| sheet | sheet.operation move | implemented | **unsupported** | `sheet.operation` move | Public position **1-based**; Office.js converts to 0-based |
| range format | range.format.read/write | implemented | **unsupported** | COM format | WPS: no in-repo format contract |
| table | list/create/delete | implemented | **unsupported** | table actions | WPS: no in-repo ListObjects contract; delete = hard delete |
| table | unlist (convertToRange) | implemented | **unsupported** | table unlist boolean | Office.js `Table.convertToRange` **ExcelApi 1.2**; keeps cell data; absence check after sync; **no** getItemOrNullObject(1.4); WPS: no ListObjects/Unlist/convertToRange contract |
| table | update (name/style/headers/totals/filter button/banded rows+columns/resize) | implemented | **unsupported** | manageWorkbookObject | Office.js Table fields; banded rows/columns **ExcelApi 1.3**; same-sheet single-area A1 resize **ExcelApi 1.13** with host overlap/header-row rules; **unsupported**: table filter rules/sort/highlight first-or-last column |
| chart | list/create/delete (column/line/bar/area/pie/scatter/doughnut/bubble/radar/linemarkers) | implemented | **unsupported** | chart actions | Office.js ChartType map via `ChartCollection.add` (**ExcelApi 1.1**); WPS: no ChartObjects contract |
| chart | update (name/type/title/style/legend/pos/size) | implemented | **unsupported** | formatChart/manage | shallow style+legend.visible; type ∈ 10 labels; `Chart.chartType` **ExcelApi 1.7**; **unsupported**: axes/data labels/source replacement/export/complex layout/stacked/3D/stock/funnel |
| chart | series list/update (name/chartType/smooth) | implemented | **unsupported** | formatChart series | Office.js `Chart.series` only name/chartType/smooth; public index 1-based; chartType 10 labels (`ChartSeries.chartType` **ExcelApi 1.7**); does **not** handle dataLabels (use chart.series.dataLabels.update); delete/add via dedicated tools; **unsupported**: formula/values/xValues/categoryFormula |

| chart | source update (setData) | implemented | **unsupported** | formatChart source | Office.js `Chart.setData` same-sheet A1 + seriesBy auto\|rows\|columns; real series readback; **unsupported**: cross-sheet source |
| chart | axes update (category/value) | implemented | **unsupported** | formatChart axes | Office.js `Chart.axes.getItem` title/min/max/majorUnit/numberFormat/reverse; primary\|secondary; **unsupported**: seriesAxis/displayUnit/log/gridlines/format |
| chart | series dataLabels update | implemented | **unsupported** | formatChart dataLabels | enabled-only: `hasDataLabels` **ExcelApi 1.7** (no dataLabels access; result omits show*); show*/numberFormat path: `ChartSeries.dataLabels` **ExcelApi 1.8** full snapshot; 1-based seriesIndex; host readback; `enabled=false` not combinable with other label fields; **unsupported**: showPercentage/showBubbleSize/delete/position/format/leaderLines |
| chart | series axisGroup update | implemented | **unsupported** | formatChart series | Office.js `ChartSeries.axisGroup` primary\|secondary only; 1-based seriesIndex; **unsupported**: formula/values/xValues/categoryFormula/add/format/trendlines |
| chart | series delete | implemented | **unsupported** | formatChart series | Office.js `ChartSeries.delete` ExcelApi 1.7; 1-based seriesIndex; delete→sync→load remaining→sync; **unsupported**: formula/values/xValues/categoryFormula/format/trendlines |
| chart | series add (empty) | implemented | **unsupported** | formatChart series | Office.js `ChartSeriesCollection.add(name?)` ExcelApi 1.7 append only; empty series `dataBound:false` (invisible until values/xValues); **unsupported**: index arg/formula |
| chart | series values/xValues update | implemented | **unsupported** | formatChart series | Office.js `setValues`/`setXAxisValues` (1.7) + `getDimensionDataSourceString` (**ExcelApi 1.15**); same-sheet A1; host source string readback; **unsupported**: formula/categoryFormula/export/cross-sheet/array literals; bubble sizes via dedicated tool |
| chart | series bubbleSizes update | implemented | **unsupported** | formatChart series bubble | Office.js `setBubbleSizes` (1.7) + `getDimensionDataSourceString("BubbleSizes")` (**ExcelApi 1.15**); same-sheet A1; bubble chart only; host source string readback; bubble type via create/update (Phase27); **unsupported**: formula/categoryFormula/array/cross-sheet/trendlines/bubbleScale/showBubbleSize/getDimensionValues-as-primary/PDF/path export |
| chart | image get (Base64) | implemented | **unsupported** | formatChart export PNG path | Office.js `Chart.getImage` **ExcelApi 1.2** → memory Base64 only; optional width/height 1–4096; **unsupported**: path write/PDF/MIME claim/fittingMode |
| range | image get (Base64 PNG) | implemented | **unsupported** | office.action snapshot range PNG | Office.js `Range.getImage` **ExcelApi 1.7** → memory Base64 only; host sheetName/address readback; **unsupported**: path write/PDF/MIME claim/width/height |




| workbook | inspect (name/active/usedRange + per-sheet dims) | implemented | implemented* | `workbook.inspect` | Office.js per-sheet address/rowCount/columnCount; empty=null/0/0. *WPS: active UsedRange address only; sheet dims not verified |
| workbook | open/create/save/switch | unsupported | unsupported | `workbook.*` | out of scope |
| macro | detect / write / run | **unsupported** | **unsupported** | `macro.*` (desktop bridge only) | 本加载项 Office.js 与 WPS 均无宏工具；desktop-source 仅记录桌面 localhost 桥，不得在加载项 runtime 宣称可用 |
| Power Query | create/manage/inspect | unsupported | unsupported | PQ ops | not claimed available |
| pivot | create/refresh/slicer | unsupported | unsupported | pivot tools | not claimed available |
| chart advanced | formula/categoryFormula/trendlines/export/complex layout | unsupported | unsupported | formatChart | Phase27 chart types include doughnut/bubble/radar/linemarkers; stacked/3D/stock/funnel not claimed; Phase26 dataLabels has enabled+four show/format fields; showPercentage/showBubbleSize/delete/position/format/leaderLines not claimed |
| file-level xlsx | Open XML | unsupported | unsupported | C# Open XML | out of scope |
| conditional format | list/add/delete (cellValue/custom) | implemented | **unsupported** | format actions | Office.js Range.conditionalFormats |
| data validation | read/write/clear (list/wholeNumber) | implemented | **unsupported** | validation | Office.js Range.dataValidation |
| COM / .NET / Electron | any | unsupported | unsupported | desktop runtime | **forbidden** |
| UI | task-pane controls beyond demo | unsupported | unsupported | N/A | task pane is host chrome only |
| UI | UserForm / ActiveX / VBA forms | unsupported | unsupported | partial desktop | **out of scope** |
| UI | custom ribbon menus / command bars | unsupported | unsupported | partial desktop | **out of scope** |
| sheet | visibility get/set (visible/hidden/veryHidden) | implemented | **unsupported** | hide/veryHide/show | Office.js `Worksheet.visibility` |
| sheet | display get/set (tabColor/showGridlines/showHeadings) | implemented | **unsupported** | worksheet update tabColor/gridlines | Office.js `Worksheet.tabColor|showGridlines|showHeadings`; empty tabColor=auto |
| sheet | freeze get/set (rows/columns/at/clear) | implemented | **unsupported** | template freezeRows | Office.js `Worksheet.freezePanes`; location writeback |
| sheet | pageLayout get/set (print settings subset) | implemented | **unsupported** | inspectPrintSettings/configurePrint | Office.js **ExcelApi 1.9**: paperSize/zoom/fit; draftMode/printOrder/firstPageNumber; margins header/footer points; headers/footers default text; **manual** horizontal/vertical page breaks (A1 append + clearPageBreaks); **unsupported**: auto page breaks/clear titles/first|even|odd/state/images/fitToOnePage |
| sheet | protection get/protect/unprotect | implemented | **unsupported** | protect/unprotect | Office.js `Worksheet.protection`; password request-memory only |
| named range | list/create/update/delete | implemented | **unsupported** | name.* | Office.js names；rename=add-then-delete（add 失败保留旧名；冲突大小写不敏感） |
| shape | list/create/delete/update (MVP) | implemented | **unsupported** | manageWorkbookObject shape | Office.js ExcelApi 1.9 `Worksheet.shapes`；geometric whitelist + textBox；shallow pos/size/text/visible；**no** image/line/group/fill/lineFormat/rotation/zOrder |

## Runnable tools

- Phase1: `host.status`, `selection.get`, `range.read/write/clear`, `formula.read/write`, `sheet.list/add/rename/delete`
- Phase3: `range.format.read/write`, `table.list/create/delete`, `chart.list/create/delete`, `workbook.inspect` (Office.js full; WPS format/table/chart unsupported)
- Phase4: `range.read` expand, `formula.context`, `sheet.operation` (add/rename/delete/copy/move)
- Phase5: `conditionalFormat.list/add/delete`, `dataValidation.read/write/clear` (Office.js via `rule`; WPS unsupported)
- Phase6: `sheet.visibility.get/set`, `sheet.protection.get/protect/unprotect`, `namedRange.list/create/update/delete` (Office.js; WPS unsupported)
- Phase7: `table.update`, `chart.update` (initial shallow fields; Office.js; WPS unsupported)
- Phase9: `sheet.display.get/set` (tabColor empty=auto or #RRGGBB, showGridlines, showHeadings; Office.js; WPS unsupported)
- Phase10: `workbook.inspect` per-sheet `usedRangeAddress`/`rowCount`/`columnCount` (Office.js; WPS dims unset)
- Phase11: `sheet.freeze.get/set` (rows|columns|at|clear; Office.js freezePanes; WPS unsupported)
- Phase12: `sheet.pageLayout.get/set` (orientation/margins/printArea/titles/zoom/flags; Office.js; WPS unsupported)
- Phase13: chart types expanded to column|line|bar|area|pie|scatter; `toChartTypeLabel` order fixed
- Phase14: chart.update/list style + legendVisible (Office.js Chart.style / legend.visible)
- Phase15: `shape.list/create/delete/update` (Office.js shapes MVP; WPS unsupported)
- Phase16: `chart.series.list` / `chart.series.update` (name/chartType/smooth only; 1-based index; Office.js; WPS unsupported)
- Phase17: `chart.source.update` (same-sheet A1 setData + seriesBy auto|rows|columns; series readback; Office.js; WPS unsupported)
- Phase18: `chart.axes.update` (category|value primary|secondary; title/min/max/majorUnit/numberFormat/reverse; Office.js; WPS unsupported)
- Phase19: `chart.series.dataLabels.update` (showValue/showCategoryName/showSeriesName/numberFormat; ExcelApi 1.8 path; 1-based seriesIndex; Office.js; WPS unsupported)
- Phase20: `chart.series.axisGroup.update` (primary|secondary; 1-based seriesIndex; Office.js; WPS unsupported)
- Phase21: `chart.series.delete` (1-based seriesIndex; remaining series readback; Office.js; WPS unsupported)
- Phase22: `chart.series.add` (optional name; empty series dataBound=false; Office.js; WPS unsupported)
- Phase23: `chart.series.values.update` (valuesRange/xValuesRange same-sheet A1; ExcelApi 1.15 source string readback; Office.js; WPS unsupported)
- Phase24: `chart.image.get` (ExcelApi 1.2 Chart.getImage Base64; optional width/height; Office.js; WPS unsupported)
- Phase25: `chart.series.bubbleSizes.update` (bubbleSizesRange same-sheet A1; setBubbleSizes + ExcelApi 1.15 BubbleSizes source string readback; bubble chart only; Office.js; WPS unsupported)
- Phase26: `chart.series.dataLabels.update` adds `enabled` via `hasDataLabels` (**ExcelApi 1.7** enabled-only path, no dataLabels access); show*/numberFormat still **ExcelApi 1.8** full snapshot; enabled=false alone; WPS unsupported
- Phase27: chart types expanded to column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers; create via ChartCollection.add enum **ExcelApi 1.1**; chart/series chartType property update **ExcelApi 1.7**; LineMarkers→linemarkers (not line), Doughnut→doughnut (not pie); WPS unsupported
- Phase28: `table.unlist` (ExcelApi 1.2 `Table.convertToRange`; keep data; host absence check; Office.js; WPS unsupported). `table.delete` remains hard delete.
- Phase29: `sheet.pageLayout` adds `paperSize` + `fitToPagesWide`/`fitToPagesTall` (**ExcelApi 1.9** precheck; host readback; fit mutually exclusive with zoomScale; WPS unsupported)
- Phase30: `sheet.pageLayout` adds `draft`/`pageOrder`/`firstPageNumber` (**ExcelApi 1.9** draftMode/printOrder/firstPageNumber; host readback; firstPageNumber ""/null→null, set only finite int≥1; WPS unsupported)
- Phase31: `sheet.pageLayout.margins` adds `header`/`footer` (**ExcelApi 1.9** headerMargin/footerMargin points; host readback; 0 allowed; WPS unsupported)
- Phase32: `sheet.pageLayout` adds default-page `headers`/`footers` left|center|right text (**ExcelApi 1.9** headersFooters.defaultForAllPages; "" clears; host readback; WPS unsupported)
- Phase33: `sheet.pageLayout` adds manual `horizontalPageBreaks`/`verticalPageBreaks` + `clearPageBreaks` (**ExcelApi 1.9** Worksheet page break collections; bare A1; append not replace; [] no-op; WPS unsupported). **unsupported**: automatic page breaks, single-break delete tool, printArea/titles clear, fitToOnePage
- Phase34: `range.image.get` (ExcelApi 1.7 Range.getImage Base64 PNG; memory only; no width/height/path/PDF/MIME; Office.js; WPS unsupported)
- Phase37: `table.update` adds `resizeAddress` (same-sheet single-area A1; ExcelApi 1.13) and `showBandedRows`/`showBandedColumns` (ExcelApi 1.3), with requirement-set precheck and write→sync→load→sync host readback; overlap/header-row geometry remains host-validated; table filter rules/sort/highlight first-or-last column remain unsupported; WPS unsupported

## Phase5 Office.js contract notes

- Conditional format custom: set `cf.custom.rule.formula` (string), not nested `formula.formula`.
- Data validation: load/set `type` + `rule` + `ignoreBlanks` only; no top-level `operator`/`formula1`/`formula2`.
- list → `rule.list.{source,inCellDropDown}`; wholeNumber → `rule.wholeNumber.{operator,formula1,formula2}`.
- Tool schema rejects unimplemented `showError`/`errorMessage` (errorAlert not wired).

## Assumptions (WPS)

In-repo verified JSA surface: Application / ActiveWorkbook name / JSIDE CodeModule via desktop bridge.
Range value/formula/sheet ops use common ET assumptions with member checks.
Format / ListObjects / ChartObjects / expand / sheet copy-move are **not** verified — typed `unsupported`.

## 交付/侧载状态

| 项 | 状态 | 说明 |
|---|---|---|
| 代码 + Vitest | 通过门槛 | `npm test` / `typecheck` / `build` 在 CI/Linux 可跑 |
| Office manifest 模板/校验 | 已提供 | `npm run manifest:dev|prod|check`；dev 默认 `https://localhost:3000` |
| 开发 HTTPS 证书工具链 | 已接入 | `office-addin-dev-certs` + `npm run certs:*` / `npm run dev` |
| 图标真实像素尺寸 | 已提供 | `public/assets/icon-{16,32,64,80}.png` IHDR 与文件名一致 |
| Windows 信任开发 CA | **未在本仓库验收** | 需在开发机执行 `certs:install` |
| Microsoft Excel 真实侧载 | **未验收** | 代码解除阻塞，不宣称已在 Excel 通过 |
| WPS 真实侧载 / jsaddons 包 | **未做** | 下一阶段 |
