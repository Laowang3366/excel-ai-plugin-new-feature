# Excel capability matrix (add-in validation)

Evidence columns:

- **office-js**: implemented against Office.js `Excel.run` / RequestContext in this package
- **wps-jsa**: only members with in-repo or strictly checked contracts; otherwise typed unsupported
- **desktop-source**: current Electron + .NET Worker tool surface (not used at runtime here)

Footnote **`implemented*`** (wps-jsa): COM-parity **member-probe** + in-repo mock unit tests. It is **not** an official WPS JSA contract and **not** real-device sideload verification. Missing members must surface as typed `unsupported` (never fake success / never COM·.NET·Shell).

| Capability family | Capability | office-js | wps-jsa | desktop-source | Evidence / notes |
|---|---|---|---|---|---|
| host | connection status | implemented | implemented | `office.connection.status` | `shared/host/*` |
| selection | selection.get | implemented | implemented* | `selection.get` | *Assumes `Application.Selection` |
| range | range.read/write/clear | implemented | implemented* | `range.*` | *Assumes Value2/Formula/Clear |
| range | range.read expand currentRegion | implemented | implemented* | expand modes | Office.js `getSurroundingRegion`; WPS `wpsJsaRangeRead` CurrentRegion member-probe + `tests/wpsJsaBasics.test.ts`; *not official JSA / not real sideload |
| range | range.read expand spill / currentArray | implemented | **unsupported** | expand modes | Office.js spill/currentArray; omit expand on single cell → spill (desktop parity). WPS typed unsupported for spill/currentArray |
| formula | formula.read / formula.write | implemented | implemented* | formula tools | writeFormulas primitive |
| formula | formula.context | implemented | implemented* | `formula.context` | Desktop shape `{sheetName,address,formulas[]}`; range optional → UsedRange |
| formula | protection inspect/manage | implemented | **unsupported** | `inspectFormulaProtection` / `manageFormulaProtection` | Office.js `Range.formulas` + `Range.format.protection.locked` **ExcelApi 1.2**; lock/unlock **formula cells only** (not whole-sheet fake); scope workbook\|sheet\|target; lock defaults unlockInputs=true (target range only) + protectSheet=true; password request-memory only, never in result; write-back verify; WPS typed unsupported; **not** real Excel sideload verified |
| formula | dependencies.inspect | implemented | implemented* | `inspectFormulaDependencies` | Pure-core text-parse graph via `shared/formulaGovernance`; scope workbook\|sheet\|target; always reports limitations (no engine circular refs); *WPS member-probe Formula/UsedRange; **not** real sideload verified |
| formula | references.repair | implemented | implemented* | `repairFormulaReferences` | Explicit mapping only; plan-first — still `#REF!` → `formula_repair_incomplete` without write; hidden `_WenggeFormulaBackup*` + `WENGGE_FORMULA_BACKUP_V1` before write; post-write verify; *WPS needs Worksheets.Add; **not** real sideload verified |
| formula | convertToValues | implemented | implemented* | `convertFormulasToValues` | Mandatory persistent backup then replace formulas with calculated values; createBackup=false rejected; *WPS member-probe; **not** real sideload verified |
| formula | backups.inspect / restore | implemented | implemented* | `inspectFormulaBackups` / `restoreFormulas` | Magic/header check; skip corrupt rows; restore by backupId with write-back verify; default retain backup (removeAfterRestore optional); *WPS member-probe; **not** real sideload verified |
| sheet | list/add/rename/delete | implemented | implemented* | `sheet.operation` | |
| sheet | sheet.operation copy | implemented | implemented* | `sheet.operation` copy | Office.js `Worksheet.copy`; WPS `wpsJsaSheetOps.wpsCopySheet` member-probe + mock tests; *not official JSA / not real sideload |
| sheet | sheet.operation move | implemented | implemented* | `sheet.operation` move | Public position **1-based**; Office.js 0-based convert; WPS `wpsJsaSheetOps.wpsMoveSheet` member-probe + mock tests; *not official JSA / not real sideload |
| range format | range.format.read/write | implemented | implemented* | COM format | Office.js Range format; WPS `wpsJsaFormat` Font/Interior/NumberFormat/Align member-probe + mock tests; *not official JSA / not real sideload |
| table | list/create/delete | implemented | **unsupported** | table actions | WPS: no in-repo ListObjects contract; delete = hard delete |
| table | unlist (convertToRange) | implemented | **unsupported** | table unlist boolean | Office.js `Table.convertToRange` **ExcelApi 1.2**; keeps cell data; absence check after sync; **no** getItemOrNullObject(1.4); WPS: no ListObjects/Unlist/convertToRange contract |
| table | update (name/style/headers/totals/filter button/banded rows+columns/first+last column/resize) | implemented | **unsupported** | manageWorkbookObject | Office.js Table fields; banded + showFirstColumn/showLastColumn **ExcelApi 1.3**; same-sheet single-area A1 resize **ExcelApi 1.13** with host overlap/header-row rules |
| table | filter get/apply/clear | implemented | **unsupported** | (desktop object schema has no filter rules) | Office.js `Table.autoFilter` apply/clearCriteria **ExcelApi 1.2**; enabled readback **ExcelApi 1.9**; public columnIndex **1-based**; filterOn values\|custom\|top/bottom items/percent; **unsupported**: cellColor/fontColor/icon/dynamic criteria detail dump; WPS unsupported; **not** real Excel sideload verified |
| table | sort get/apply/clear | implemented | **unsupported** | (desktop object schema has no table sort) | Office.js `Table.sort` apply/clear/fields **ExcelApi 1.2**; fields ≤3; public columnIndex **1-based**; value sort only; color/icon sort **unsupported**; WPS unsupported; **not** real Excel sideload verified |
| chart | list/create/delete (column/line/bar/area/pie/scatter/doughnut/bubble/radar/linemarkers) | implemented | **unsupported** | chart actions | Office.js ChartType map via `ChartCollection.add` (**ExcelApi 1.1**); WPS: no ChartObjects contract |
| chart | update (name/type/title/style/legend/pos/size) | implemented | **unsupported** | formatChart/manage | shallow style+legend.visible; type ∈ 10 labels; `Chart.chartType` **ExcelApi 1.7**; **unsupported**: axes/data labels/source replacement/export/complex layout/stacked/3D/stock/funnel |
| chart | series list/update (name/chartType/smooth) | implemented | **unsupported** | formatChart series | Office.js `Chart.series` only name/chartType/smooth; public index 1-based; chartType 10 labels (`ChartSeries.chartType` **ExcelApi 1.7**); does **not** handle dataLabels (use chart.series.dataLabels.update); delete/add via dedicated tools; **unsupported**: formula/values/xValues/categoryFormula |

| chart | source update (setData) | implemented | **unsupported** | formatChart source | Office.js `Chart.setData(Range)` same-sheet **or** cross-sheet A1 (`Sheet2!A1` / `'Sheet 2'!A1`) + seriesBy auto\|rows\|columns; real series readback; **unsupported**: external workbook / 3D / multi-area / structured refs; WPS unsupported; **not** real sideload verified |
| chart | axes update (category/value) | implemented | **unsupported** | formatChart axes | Office.js `Chart.axes.getItem` title/min/max/majorUnit|minorUnit(`""`=auto)/numberFormat/reverse + **ExcelApi 1.7** displayUnit/customDisplayUnit(`setCustomDisplayUnit`)/scaleType/logBase/showDisplayUnitLabel/majorTickMark|minorTickMark|tickLabelPosition + **ExcelApi 1.8** position/`setPositionAt`(positionAt) + **ExcelApi 1.9** linkNumberFormat + **ExcelApi 1.1** major/minorGridlines.visible; primary|secondary; write→sync→load→sync; **unsupported**: seriesAxis/axis format/trendline labels; WPS unsupported; **not** real sideload verified |
| chart | series dataLabels update | implemented | **unsupported** | formatChart dataLabels | enabled-only: `hasDataLabels` **ExcelApi 1.7** (no dataLabels access; result omits show*); showValue/showCategoryName/showSeriesName/numberFormat/**showPercentage/showBubbleSize/showLegendKey/separator/position** path: `ChartSeries.dataLabels` **ExcelApi 1.8** full snapshot (position enum w/o Invalid; separator preserved as-is); 1-based seriesIndex; host readback; `enabled=false` not combinable with other label fields; **unsupported**: delete/format/leaderLines; WPS unsupported; **not** real sideload verified |
| chart | series axisGroup update | implemented | **unsupported** | formatChart series | Office.js `ChartSeries.axisGroup` primary\|secondary only; 1-based seriesIndex; **unsupported**: formula/values/xValues/categoryFormula/add/format/trendlines |
| chart | series delete | implemented | **unsupported** | formatChart series | Office.js `ChartSeries.delete` ExcelApi 1.7; 1-based seriesIndex; delete→sync→load remaining→sync; **unsupported**: formula/values/xValues/categoryFormula/format/trendlines |
| chart | series add (empty) | implemented | **unsupported** | formatChart series | Office.js `ChartSeriesCollection.add(name?)` ExcelApi 1.7 append only; empty series `dataBound:false` (invisible until values/xValues); **unsupported**: index arg/formula |
| chart | series values/xValues update | implemented | **unsupported** | formatChart series | Office.js `setValues`/`setXAxisValues` (1.7) + `getDimensionDataSourceString` (**ExcelApi 1.15**); same-sheet A1; host source string readback; **unsupported**: formula/categoryFormula/export/cross-sheet/array literals; bubble sizes via dedicated tool |
| chart | series bubbleSizes update | implemented | **unsupported** | formatChart series bubble | Office.js `setBubbleSizes` (1.7) + `getDimensionDataSourceString("BubbleSizes")` (**ExcelApi 1.15**); same-sheet A1; bubble chart only; host source string readback; bubble type via create/update (Phase27); **unsupported**: formula/categoryFormula/array/cross-sheet/trendlines/bubbleScale/getDimensionValues-as-primary/PDF/path export; this tool does not handle label visibility — use `chart.series.dataLabels.update` for showBubbleSize |
| chart | series trendlines list/add/update/delete | implemented | **unsupported** | formatChart trendlines | Office.js `ChartSeries.trendlines` **ExcelApi 1.7** (`add`/`getItem`(0-based host)/`delete`, type/name/intercept/polynomialOrder/movingAveragePeriod); intercept write `""`=automatic, readback number; public trendlineIndex 1-based; **ExcelApi 1.8** forward/backwardPeriod/showEquation/showRSquared; write→sync→load returned object+getCount; **unsupported**: format/label styling/`name=null` auto, formula; WPS unsupported; **not** real sideload verified |
| chart | series trendlines format update | implemented | **unsupported** | formatChart trendline format | Office.js `ChartTrendline.format.line` **ExcelApi 1.7** (`color` HTML/#RRGGBB repo convention, `lineStyle` full `ChartLineStyle` enum, `weight` finite points — **no official min/max**); 1-based seriesIndex+trendlineIndex; write→sync→load→sync; **unsupported**: label styling/font/fill, clear(); WPS unsupported; **not** real sideload verified |
| chart | series markers update | implemented | **unsupported** | formatChart series markers | Office.js `ChartSeries.markerStyle/markerSize/markerBackgroundColor/markerForegroundColor` **ExcelApi 1.7**; style automatic\|none\|square\|diamond\|triangle\|x\|star\|dot\|dash\|circle\|plus\|picture (no Invalid write); size 2–72; colors **#RRGGBB**; 1-based seriesIndex; write→sync→load→sync; **unsupported**: per-point markers/picture content; WPS unsupported; **not** real sideload verified |
| chart | image get (Base64) | implemented | **unsupported** | formatChart export PNG path | Office.js `Chart.getImage` **ExcelApi 1.2** → memory Base64 only; optional width/height 1–4096; **unsupported**: path write/PDF/MIME claim/fittingMode |
| range | image get (Base64 PNG) | implemented | **unsupported** | office.action snapshot range PNG | Office.js `Range.getImage` **ExcelApi 1.7** → memory Base64 only; host sheetName/address readback; **unsupported**: path write/PDF/MIME claim/width/height |
| range | insert / delete | implemented | implemented* | row/column insert/delete | Office.js `Range.insert/delete` **ExcelApi 1.1** shift down\|right / up\|left; *WPS: `Range.Insert`/`Delete` + xlShift member-probe + mock tests; **not** official JSA / **not** real sideload |
| range | autofit | implemented | implemented* | row/column autofit | Office.js `RangeFormat.autofitRows` / `RangeFormat.autofitColumns` **ExcelApi 1.2** with rowHeight/columnWidth readback; WPS `wpsAutofitRange` member-probe + mock tests; *not official JSA / not real sideload |




| workbook | inspect (name/active/usedRange + per-sheet dims) | implemented | implemented* | `workbook.inspect` | Office.js per-sheet address/rowCount/columnCount; empty=null/0/0. *WPS: active UsedRange address only; sheet dims not verified |
| workbook | objects.inspect (sheets + tables/charts/namedRanges/shapes) | implemented | implemented* | `workbook.objects.inspect` | Safe inventory; categories status available\|unsupported\|failed; maxItemsPerCategory 1..500 default 100; truncated keeps totalCount; partial category failure OK. Office.js batched; *WPS: sheets+namedRanges available*, table/chart/shape unsupported (not empty success); **not** real sideload verified |
| workbook | save (current workbook in place) | implemented | implemented* | `workbook.save` | Office.js `Workbook.save` **ExcelApi 1.1** (no path/saveAs); *WPS: `ActiveWorkbook.Save` member-probe + bridge evidence (`desktop/public/wps-jsa-bridge`); unnamed book may host-dialog/fail; **not** real sideload verified |
| workbook | open/create/saveAs/switch (path) | unsupported | unsupported | `workbook.*` | out of scope; path/disk lifecycle not claimed |
| macro | detect / write / run | **unsupported** | **unsupported** | `macro.*` (desktop bridge only) | 本加载项 Office.js 与 WPS 均无宏工具；desktop-source 仅记录桌面 localhost 桥，不得在加载项 runtime 宣称可用 |
| Power Query | create/manage/inspect | unsupported | unsupported | PQ ops | not claimed available |
| pivot | list/create/refresh (current workbook) | implemented | **unsupported** | `pivot.list` / `pivot.create` / `pivot.refresh` | Office.js list/create **ExcelApi 1.8** (`pivotTables.add` + hierarchies); refresh **ExcelApi 1.3** (`PivotTable.refresh`); `refreshConnections=true` → **ExcelApi 1.7** `Workbook.dataConnections.refreshAll` (**limited** supported connections only — e.g. Pivot→Power BI, same-workbook Data Model→table/range; **not** Power Query / external workbook except Power BI / firewall; **no** connection status readback → `verified:false`; **not** full desktop `Workbook.RefreshAll` parity). ≥1 field required on create; dataFields may multi-agg; empty/`""` destination → `Pivots` lastBottom+3. slicers → separate `slicer.*` tools (ExcelApi 1.10); WPS pivot typed unsupported; **not** real Excel sideload verified |
| slicer | list/create/update/delete + filter get/apply/clear | implemented | **unsupported** | `slicer.list` / `slicer.create` / `slicer.update` / `slicer.delete` / `slicer.filter.get` / `slicer.filter.apply` / `slicer.filter.clear` | Office.js **ExcelApi 1.10** stable subset only (`caption/height/left/name/sortBy/style/top/width` R/W; `id/isFilterCleared/worksheet/slicerItems` RO; `clearFilters/delete/getSelectedItems/selectItems`). `selectItems([])` = select **all** (not none). **No** stable source readback → create returns `requestedSource` + limitations. BETA `nameInFormula`/`slicerStyle`/`setStyle` **not** exposed. create requires `advancedIntent=interactive-pivot`, `sourceType=table|pivotTable`. strict scalar/item/selectedKeys readback (no coercion); select-all/clear verify all items + isFilterCleared; missing sheet list fails; WPS typed unsupported (no JSA evidence); **not** real Excel/WPS sideload verified |
| chart advanced | formula/categoryFormula/export/complex layout | unsupported | unsupported | formatChart | Phase47 trendlines + Phase48 series markers + Phase49 trendline line format; formula/categoryFormula still no Office.js contract; stacked/3D/stock/funnel not claimed; delete/format/leaderLines not claimed (showPercentage/showBubbleSize/showLegendKey/separator/position via dataLabels.update Phase52) |
| file-level xlsx | Open XML | unsupported | unsupported | C# Open XML | out of scope |
| conditional format | list/add/delete (cellValue/custom; list honest hostType) | implemented | implemented* | format actions | Office.js **ExcelApi 1.6** `Range.conditionalFormats`; list keeps real hostType (ContainsText/DataBar/…); add only cellValue/custom; *WPS: FormatConditions 1-based index ids + mock; **not** real sideload verified |
| data validation | read/write/clear (list/wholeNumber/decimal/date/time/textLength/custom + errorAlert/prompt) | implemented | implemented* | validation | Office.js **ExcelApi 1.8** `Range.dataValidation` rule + `ignoreBlanks`←`allowBlank` + **errorAlert**(showAlert/title/message/style) + **prompt**(showPrompt/title/message); list source = inline string or Range proxy; Inconsistent/MixedCriteria → rule=null + limitations; omit errorAlert/prompt on write preserves host metadata; *WPS: Validation rule/allowBlank snapshot+restore; errorAlert/prompt **typed unsupported** (no JSA member evidence); **not** real sideload verified |
| COM / .NET / Electron | any | unsupported | unsupported | desktop runtime | **forbidden** |
| UI | task-pane controls beyond demo | unsupported | unsupported | N/A | task pane is host chrome only |
| UI | UserForm / ActiveX / VBA forms | unsupported | unsupported | partial desktop | **out of scope** |
| UI | custom ribbon menus / command bars | unsupported | unsupported | partial desktop | **out of scope** |
| sheet | visibility get/set (visible/hidden/veryHidden) | implemented | implemented* | hide/veryHide/show | Office.js `Worksheet.visibility`; *WPS: `Worksheet.Visible` member-probe + mock only |
| sheet | display get/set (tabColor/showGridlines/showHeadings) | implemented | **unsupported** | worksheet update tabColor/gridlines | Office.js `Worksheet.tabColor|showGridlines|showHeadings`; empty tabColor=auto |
| sheet | freeze get/set (rows/columns/at/clear) | implemented | **unsupported** | template freezeRows | Office.js `Worksheet.freezePanes`; location writeback |
| sheet | pageLayout get/set (print settings subset) | implemented | **unsupported** | inspectPrintSettings/configurePrint | Office.js **ExcelApi 1.9**: orientation/margins/headers|footers/draft/pageOrder/firstPageNumber/paperSize/zoomScale; **printArea** + **printTitleRows/Columns** (set non-empty + get OrNullObject; desktop aliases **repeatRows/repeatColumns**); **fitToPagesWide/Tall** + desktop **fitToOnePageWide/Tall**→pages=1; manual page breaks; **unsupported**: printArea/titles **clear** (no official no-arg API), auto page breaks, first|even|odd headers, images; WPS unsupported; **not** real sideload |
| sheet | protection get/protect/unprotect | implemented | implemented* | protect/unprotect | Office.js `Worksheet.protection`; password request-memory only; *WPS: ProtectContents/Protect/Unprotect member-probe; password never echoed |
| named range | list/create/update/delete | implemented | implemented* | name.* | Office.js names；rename=add-then-delete；*WPS: Names Add/Delete member-probe；rename Add 先于 Delete，失败回滚 |
| shape | list/create/delete/update (MVP) | implemented | **unsupported** | manageWorkbookObject shape | Office.js ExcelApi 1.9 `Worksheet.shapes`；geometric whitelist + textBox；shallow pos/size/text/visible；**no** image/line/group/fill/lineFormat/rotation/zOrder |

## Runnable tools

- Phase1: `host.status`, `selection.get`, `range.read/write/clear`, `formula.read/write`, `sheet.list/add/rename/delete`
- Phase3: `range.format.read/write`, `table.list/create/delete`, `chart.list/create/delete`, `workbook.inspect` (Office.js full; **WPS format = implemented* member-probe**; table/chart still unsupported)
- Phase4: `range.read` expand, `formula.context`, `sheet.operation` (add/rename/delete/copy/move). WPS: currentRegion + copy/move = implemented* (member-probe); spill/currentArray still unsupported
- Phase54: `slicer.*` Office.js ExcelApi 1.10 stable subset (no source readback; selectItems([])=all; WPS unsupported). Phase5/53: `conditionalFormat.list/add/delete` (ExcelApi 1.6; list honest hostType; add cellValue/custom only), `dataValidation.read/write/clear` (ExcelApi 1.8; list/wholeNumber/decimal/date/time/textLength/custom; `allowBlank`↔`ignoreBlanks`; **errorAlert/prompt** Office.js top-level metadata with write→sync→load→sync; Inconsistent/MixedCriteria honest); WPS CF/DV rules **implemented*** via FormatConditions/Validation member-probe; WPS errorAlert/prompt **unsupported**; **not** real Excel/WPS sideload verified
- Phase6: `sheet.visibility.get/set`, `sheet.protection.get/protect/unprotect`, `namedRange.list/create/update/delete` (Office.js + **WPS implemented*** member-probe; not device-verified)
- Phase7: `table.update`, `chart.update` (initial shallow fields; Office.js; WPS unsupported)
- Phase9: `sheet.display.get/set` (tabColor empty=auto or #RRGGBB, showGridlines, showHeadings; Office.js; WPS unsupported)
- Phase10: `workbook.inspect` per-sheet `usedRangeAddress`/`rowCount`/`columnCount` (Office.js; WPS dims unset)
- Phase44: `sheet.pageLayout` printArea/printTitle* (Office.js set+get; aliases repeatRows/Columns; fitToOnePageWide/Tall→fit pages 1; clear still unsupported)
- Phase43: `workbook.save` current workbook only (Office.js ExcelApi 1.1; WPS Save member-probe; no open/create/saveAs/switch)
- Phase42: `workbook.objects.inspect` capped object inventory (Office.js full categories; WPS sheets+names; table/chart/shape unsupported categories)
- Phase11: `sheet.freeze.get/set` (rows|columns|at|clear; Office.js freezePanes; WPS unsupported)
- Phase12: `sheet.pageLayout.get/set` (orientation/margins/printArea/titles/zoom/flags; Office.js; WPS unsupported)
- Phase13: chart types expanded to column|line|bar|area|pie|scatter; `toChartTypeLabel` order fixed
- Phase14: chart.update/list style + legendVisible (Office.js Chart.style / legend.visible)
- Phase15: `shape.list/create/delete/update` (Office.js shapes MVP; WPS unsupported)
- Phase16: `chart.series.list` / `chart.series.update` (name/chartType/smooth only; 1-based index; Office.js; WPS unsupported)
- Phase17: `chart.source.update` (Office.js `Chart.setData(Range)`: same-sheet **or** cross-sheet `Sheet2!A1` / `'Sheet 2'!A1`; rejects external/3D/multi-area/structured; series readback; WPS unsupported)
- Phase18/46/50: `chart.axes.update` (category|value primary|secondary; title/min/max/majorUnit|minorUnit(`""` auto)/numberFormat/reverse; Phase46: displayUnit/customDisplayUnit/scaleType/logBase/showDisplayUnitLabel + gridlines; Phase50: major/minorTickMark/tickLabelPosition ExcelApi 1.7, position/setPositionAt ExcelApi 1.8, linkNumberFormat ExcelApi 1.9; Office.js; WPS unsupported; **not** real sideload verified)
- Phase19/26/52: `chart.series.dataLabels.update` (enabled ExcelApi 1.7; show*/numberFormat/showPercentage/showBubbleSize/showLegendKey/separator/position ExcelApi 1.8 full snapshot; position w/o Invalid; Office.js; WPS unsupported; **not** real sideload verified)
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
- Phase33: `sheet.pageLayout` adds manual `horizontalPageBreaks`/`verticalPageBreaks` + `clearPageBreaks` (**ExcelApi 1.9** Worksheet page break collections; bare A1; append not replace; [] no-op; WPS unsupported). **unsupported**: automatic page breaks, single-break delete tool, printArea/titles **clear** (no official clear API); fitToOnePage covered via fitToOnePageWide/Tall aliases (Phase44)
- Phase34: `range.image.get` (ExcelApi 1.7 Range.getImage Base64 PNG; memory only; no width/height/path/PDF/MIME; Office.js; WPS unsupported)
- Phase38: `range.insert` / `range.delete` (ExcelApi 1.1 Office.js + **WPS implemented*** xlShift member-probe) and `range.autofit` (ExcelApi 1.2; Office.js + **WPS implemented*** member-probe)
- Phase37: `table.update` adds `resizeAddress` (same-sheet single-area A1; ExcelApi 1.13) and `showBandedRows`/`showBandedColumns` (ExcelApi 1.3), with requirement-set precheck and write→sync→load→sync host readback; overlap/header-row geometry remains host-validated; WPS unsupported
- Phase41: formula governance tools wired to pure core — `formula.dependencies.inspect` (safe), `formula.references.repair` / `formula.convertToValues` / `formula.backups.restore` (dangerous), `formula.backups.inspect` (safe); plan-first repair; mandatory backup for convert; WPS member-probe paths; **not** real sideload verified
- Phase40: `formula.protection.inspect` (safe) + `formula.protection.manage` (dangerous): ExcelApi 1.2 locked on formula cells only; scope workbook|sheet|target; unlockInputs range-limited; protectSheet optional; password request-memory only; post-write verify; WPS unsupported; **not** real Excel sideload verified
- Phase39: `table.update` adds `showFirstColumn`/`showLastColumn` (**ExcelApi 1.3**); new tools `table.filter.get|apply|clear` (AutoFilter apply/clear **1.2**, enabled **1.9**) and `table.sort.get|apply|clear` (**ExcelApi 1.2**); 1-based columnIndex; filterOn values|custom|top/bottom items/percent; sort fields≤3 value-only; color/icon/dynamic filter & color sort remain typed unsupported; WPS typed unsupported; **not** real Excel sideload verified
- Phase47: `chart.series.trendlines.list/add/update/delete` (ExcelApi 1.7/1.8; intercept `""` auto; host getItem 0-based mapped to public 1-based; Office.js; WPS unsupported; **not** real sideload verified)
- Phase48: `chart.series.markers.update` (markerStyle/size 2–72/colors #RRGGBB; ExcelApi 1.7; Office.js; WPS unsupported; **not** real sideload verified)
- Phase49: `chart.series.trendlines.format.update` (format.line color/lineStyle/weight ExcelApi 1.7; weight finite points only — no official range; Office.js; WPS unsupported; **not** real sideload verified)

## Phase5 Office.js contract notes

- CF ExcelApi **1.6**; DV ExcelApi **1.8**; precheck via `isSetSupported` (missing/throw → unsupported).
- Conditional format custom: set `cf.custom.rule.formula` (ClientObject property), not nested `formula.formula`; cellValue.rule is whole-object assign. CF host operators: GTE/LTE **without** `To`; notEqualTo → **NotEqualTo**. add: `cf.load("id,type")` before first sync; verify rule + colors after second sync.
- Data validation: load/set `type` + `rule` + `ignoreBlanks` (+ optional `errorAlert`/`prompt`). Public `rule.allowBlank` maps to host `ignoreBlanks` (no duplicate public ignoreBlanks field). No top-level `operator`/`formula1`/`formula2`.
- DV writable types: **list / wholeNumber / decimal / date / time / textLength / custom**. list source = inline string **or Excel.Range proxy** (not `String(range)`); formula/range reads keep `listSourceKind=range` (never split to listValues). Inconsistent/MixedCriteria/unknown/lossy inline → `rule:null` + limitations. clear requires hostType **None**.
- **Phase53/53.1/53.2** `errorAlert`/`prompt` are DataValidation **top-level metadata** (not rule fields). ClientObject: load top-level `type,rule,ignoreBlanks,errorAlert,prompt` → sync → **assert complete surface** (official type tokens None|WholeNumber|Decimal|List|Date|Time|TextLength|Custom|Inconsistent|MixedCriteria; boolean ignoreBlanks; object rule; full alerts) → pure **planDvRuleWrite** (XOR list sources, between formula2, reject illegal combos) → whole-object assign → write sync → reload. Empty `{}` rejected; omit metadata preserves host. Incomplete/illegal → ordinary failed zero writes. ExcelApi **1.8**. Enum readback (type/operator/alertStyle) is **case-insensitive exact only** (no trim/space-collapse). Inline listValues maxItems 1000 / ≤255 chars.
- WPS: CF/DV **implemented*** (`FormatConditions` 1-based index ids; `Validation` snapshot+restore on write; list uses shared `classifyListSource`/`dvRulesMatch`). **Not** real Excel/WPS sideload verified.

## Formula governance (wired tools + pure core)

`shared/formulaGovernance/` is the pure address/dependency/repair/backup library. **Callable tools**: `formula.dependencies.inspect`, `formula.references.repair`, `formula.convertToValues`, `formula.backups.inspect`, `formula.backups.restore`, plus `formula.protection.inspect` / `formula.protection.manage`. Backup protocol `WENGGE_FORMULA_BACKUP_V1` on hidden `_WenggeFormulaBackup*` sheets. Dependency analysis is text-parse only (not Excel calc engine). **Not** real Excel/WPS sideload verified.

## Assumptions (WPS)

In-repo verified JSA surface (desktop bridge history): Application / ActiveWorkbook name / JSIDE CodeModule.
Range value/formula/sheet ops use common ET assumptions with member checks.

**WPS implemented*** (member-probe + mock/unit tests only; **not** official JSA contract; **not** real device sideload):

- `range.read` expand **currentRegion** (`wpsJsaRangeRead.ts`)
- `sheet.operation` **copy/move** (`wpsJsaSheetOps.ts`)
- `range.format.read/write` (`wpsJsaFormat.ts`)
- `range.autofit` / `range.insert` / `range.delete` (`wpsJsaRangeStructureUnsupported.ts`; Insert/Delete use xlShift constants)
- `sheet.visibility.get/set` (`wpsJsaSheetVisibility.ts`)
- `sheet.protection.get/protect/unprotect` (`wpsJsaSheetProtection.ts`; password request-memory only, never echoed)
- `namedRange.list/create/update/delete` (`wpsJsaNamedRanges.ts`; rename = Add then Delete with rollback)
- formula governance: dependencies / repair / convertToValues / backups inspect|restore (`wpsJsaFormulaGovernance*`)
- `conditionalFormat.list/add/delete` (`wpsJsaConditionalFormat.ts`; FormatConditions 1-based index ids)
- `dataValidation.read/write/clear` (`wpsJsaDataValidation*.ts`; Validation snapshot restore on write failure)

**WPS still typed unsupported** (do not enlarge claims):

- expand **spill** / **currentArray**
- `formula.protection.*`
- table / chart / pageLayout / freeze / display / shape / range.image / chart.image
- macros / Power Query / Pivot / arbitrary-path OpenXML / workbook open-create-save-switch / cross-Office disk transactions


## WPS remaining-capability audit

See [`wps-remaining-capability-audit.md`](./wps-remaining-capability-audit.md): no additional WPS features without bridge/member evidence; package gates distinguish prompt/doc text from runtime Electron/COM/.NET/child_process imports.

## 交付/侧载状态

| 项 | 状态 | 说明 |
|---|---|---|
| 代码 + Vitest | 通过门槛 | `npm test` / `typecheck` / `build` 在 CI/Linux 可跑 |
| Office manifest 模板/校验 | 已提供 | `npm run manifest:dev|prod|check`；dev 默认 `https://localhost:3000` |
| 开发 HTTPS 证书工具链 | 已接入 | `office-addin-dev-certs` + `npm run certs:*` / `npm run dev` |
| 图标真实像素尺寸 | 已提供 | `public/assets/icon-{16,32,64,80}.png` IHDR 与文件名一致 |
| Windows 信任开发 CA | **未在本仓库验收** | 需在开发机执行 `certs:install` |
| Microsoft Excel 真实侧载 | **未验收** | 代码解除阻塞，不宣称已在 Excel 通过 |
| WPS 正式本地 jsaddons 包生成 | **可生成** | `npm run package:wps`；布局对齐桌面 bridge 的 publish/url 合同；**真实 WPS 侧载尚未验收** |
| Office 生产静态包门禁 | **已实现** | `npm run package:prod -- --base-url https://…`；拒绝 localhost/http 残留；**真实 Excel 侧载尚未验收** |
| WPS 源校验命令 | 已提供 | `npm run manifest:wps:check`（`manifest:check` 一并执行） |
