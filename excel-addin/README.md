# 文格 Excel 加载项验证项目

独立 `excel-addin/` 包：Office.js + WPS JSA 任务窗格骨架，**不**依赖 Electron / COM / .NET Worker，也不是根 workspace 成员。

## 命令

```powershell
cd excel-addin
npm ci
npm run sync:prompts   # 从 desktop 同步 Excel 相关提示词 + SHA-256 manifest
npm run typecheck
npm test
npm run build
npm run dev            # https 侧载前可用 http://localhost:3000
```

## 结构

| 路径 | 职责 |
|------|------|
| `shared/host` | HostAdapter、Office.js / WPS JSA 适配 |
| `shared/tools` | Phase 1–29 工具合同与执行（写后回读可显式开启） |
| `shared/provider` | 供应商模板 CRUD / active / apiFormat；key 仅内存 |
| `shared/prompts` | 同步生成的提示词与 manifest |
| `docs/capability-matrix.md` | 能力矩阵与证据 |
| `manifest/` | Office manifest + WPS JSA 安装说明 |
| `src/` | Task pane UI |

## 安全存储说明

API key 默认只存在 `MemorySecretStore`（进程内存）。**禁止**写入 `localStorage`。跨会话持久化需后续本地安全存储服务。public view 只暴露 `hasApiKey`，不包含密钥明文。

## Provider 连接测试 / 模型列表

窄 `ProviderClient`（`shared/provider/client.ts`）与桌面 `ai:testConnection` / `ai:listModels` 语义对齐：

| apiFormat | 测试连接 | 模型列表 |
|-----------|----------|----------|
| `openai` | `POST {base}/chat/completions` | `GET {base}/models` |
| `responses` | `POST {base}/responses` | `GET {base}/models` |
| `anthropic` | `POST {base}/messages` + `x-api-key` + `anthropic-version` | **unsupported**（无列表接口） |

**限制：**

- 任务窗格内浏览器 `fetch` 直连第三方 API 常被 **CORS** 拒绝；客户端将 `Failed to fetch` 归类为 `cors`/`network` 并原样返回，不伪造成功。
- 无 API key 时返回 `missing_key`。
- HTTP 非 2xx 返回 `http` + status + 错误消息；不做重试。
- 后续若需稳定连通，应经本地安全代理（不在本批范围）。

## 加载

- Microsoft Excel：侧载 `manifest/office-excel-manifest.xml`（开发时 SourceLocation / 图标均指向 `https://localhost:3000`，图标文件在 `public/assets/icon-16|32|64|80.png`，`npm run dev` / `build` 会由 Vite 提供；需按 Office 要求配置 HTTPS 证书）。
- WPS：见 `manifest/wps-jsa/README.md`。

## Excel 工具 parity（Phase 4–34）

| 工具 | 说明 |
|------|------|
| `range.read` | 可选 `expand: none\|spill\|currentArray\|currentRegion`。**省略 expand 且单单元格**时与桌面一致自动 spill；显式 `none` 不扩展。WPS 上非 `none` expand → unsupported |
| `formula.context` | 返回桌面兼容形状 `{ sheetName, address, formulas:[{address,formula,value}] }`。`range` 可省略→UsedRange |
| `sheet.operation` | `add\|rename\|delete\|copy\|move`。**position 为 1-based**（与桌面 COM 合同一致；Office.js 内部转 0-based）。WPS copy/move → unsupported |
| `conditionalFormat.list/add/delete` | Office.js：`cellValue`（必填 operator/formula1）与 `custom`（必填 formula 字符串）。WPS → unsupported |
| `dataValidation.read/write/clear` | Office.js：经 `DataValidation.rule`（list / wholeNumber）；**不**使用顶层 formula1/operator。wholeNumber 必填 operator。WPS → unsupported |
| `sheet.visibility.get/set` | Office.js：`visible\|hidden\|veryHidden`。WPS → unsupported |
| `sheet.protection.get/protect/unprotect` | Office.js：`Worksheet.protection`；**password 仅请求内存**，不写 localStorage/日志。WPS → unsupported |
| `namedRange.list/create/update/delete` | Office.js：`Workbook.names` / `Worksheet.names`（scope workbook\|worksheet）。`newName` 重命名：先 add 新名再 delete 旧名（add 失败则旧名保留）；冲突大小写不敏感。WPS → unsupported |
| `table.update` | Office.js 浅层：newName/style/showHeaders/showTotals/showFilterButton（≥1 字段）。**无** resize/filter/sort/banded/highlight。WPS → unsupported |
| `table.unlist` | Office.js：**ExcelApi 1.2** `Table.convertToRange` → 保留全部单元格数据并转为普通区域；宿主 sheet/table/address 回读 + absence 校验；`unlisted:true`。`table.delete` 仍为硬删除表格。WPS → unsupported |
| `chart.create` | Office.js：column\|line\|bar\|area\|pie\|scatter\|doughnut\|bubble\|radar\|linemarkers（省略 chartType→column；`ChartCollection.add` enum **ExcelApi 1.1**）。WPS → unsupported |
| `chart.update` | Office.js 浅层：newName/chartType(十种)/title/showTitle/style(正整数)/showLegend/left/top/width/height（≥1 字段；w/h>0；`Chart.chartType` **ExcelApi 1.7**）。**无** axes/data labels/source replacement/export/复杂布局/stacked/3D/stock/funnel。WPS → unsupported |
| `chart.series.list` / `chart.series.update` | Office.js：series 最小子集 name/chartType/smooth；`seriesIndex` **1-based**；update ≥1 字段；chartType 十种（`ChartSeries.chartType` **ExcelApi 1.7**）。**不处理 dataLabels**（见 `chart.series.dataLabels.update`）。删除/创建见 `chart.series.delete` / `chart.series.add`。**无** formula/values/xValues/categoryFormula。WPS → unsupported |
| `chart.source.update` | Office.js：`Chart.setData` 同表 A1 `sourceRange` + `seriesBy` auto\|rows\|columns（默认 auto）；写后 sync 并回读 series 快照。**无** 跨表源、data labels、export。WPS → unsupported |
| `chart.axes.update` | Office.js：`kind` category\|value；`group` primary\|secondary（默认 primary）；title（空串清除）/minimum/maximum/majorUnit(≥0)/numberFormat/reverse（≥1 字段）；写后真回读。**无** seriesAxis/displayUnit/log/gridlines。WPS → unsupported |
| `chart.series.dataLabels.update` | Office.js：seriesIndex **1-based**。**enabled-only** → ExcelApi **1.7** `hasDataLabels`（不访问 `dataLabels`；结果仅 sheet/chart/seriesIndex/enabled）。触碰 showValue/showCategoryName/showSeriesName/numberFormat → ExcelApi **1.8** 完整快照。`enabled=false` 不可与其它标签字段同传。**unsupported**：showPercentage/showBubbleSize/delete/position/format/leaderLines。WPS → unsupported |
| `chart.series.axisGroup.update` | Office.js：seriesIndex **1-based**；`axisGroup` primary\|secondary；写后真回读。**无** formula/values/xValues/categoryFormula/add/format/trendlines。WPS → unsupported |
| `chart.series.delete` | Office.js：seriesIndex **1-based**；`ChartSeries.delete`→sync→load remaining→sync；回读 `{ deletedSeriesIndex, remainingSeries }`（index 连续 1-based）。**无** formula/values/xValues/categoryFormula/format/trendlines。WPS → unsupported |
| `chart.series.add` | Office.js：`ChartSeriesCollection.add(name?)` 仅 append；可选 `name`；空 series `dataBound:false`（设置 values/xValues 前图中不可见）。**无** index 参数、formula。WPS → unsupported |
| `chart.series.values.update` | Office.js：**ExcelApi 1.15**；`setValues`/`setXAxisValues` + `getDimensionDataSourceString` 真源回读；`valuesRange`/`xValuesRange` 同表 A1（≥1）；`dataBound:true`。**无** formula/categoryFormula/export/跨表/数组字面量。bubble sizes 见 `chart.series.bubbleSizes.update`。WPS → unsupported |
| `chart.series.bubbleSizes.update` | Office.js：**ExcelApi 1.15** verified readback；`setBubbleSizes` + `getDimensionDataSourceString("BubbleSizes")`；`bubbleSizesRange` 同表 A1；仅 bubble chart 有效；`dataBound:true`。bubble chart type 见 create/update（Phase27）。**无** formula/categoryFormula、数组值、Categories/YValues、跨表、trendlines、bubbleScale、dataLabels.showBubbleSize、getDimensionValues 主验证、PDF/path export。WPS → unsupported |
| `chart.image.get` | Office.js：**ExcelApi 1.2** `Chart.getImage` → 内存 Base64；可选 width/height（1–4096）；宿主 chartName 回读。**无** 路径写入/PDF/fittingMode/MIME 声明。WPS → unsupported |
| `range.image.get` | Office.js：**ExcelApi 1.7** `Range.getImage` → 内存 Base64 区域 PNG；宿主 sheetName/address 真回读。**无** width/height、路径、PDF、MIME 声明。WPS → unsupported |



| `sheet.display.get/set` | Office.js：`tabColor`（空串=自动色或 `#RRGGBB`）、`showGridlines`、`showHeadings`；set ≥1 字段写后回读。WPS → unsupported |
| `workbook.inspect` | Office.js：活动表 usedRange + 每表 `usedRangeAddress`/`rowCount`/`columnCount`（空表 null/0/0）。WPS：活动表 address；每表尺寸字段不填（未验证 Rows/Columns） |
| `sheet.freeze.get/set` | Office.js：`freezePanes` rows/columns/at/clear；get 回读 location（无冻结 null/0/0）。WPS → unsupported |
| `sheet.pageLayout.get/set` | Office.js：**ExcelApi 1.9** `pageLayout` + 手动分页符（orientation/margins points/headers|footers default 六槽/print flags/zoomScale/paperSize/fit/draft/pageOrder/firstPageNumber/printArea/titles/**clearPageBreaks + horizontal|verticalPageBreaks bare A1 append**）；fit 与 zoomScale 互斥。**无** clear titles/自动分页/first|even|odd/headersFooters state/图片/fitToOnePage。WPS → unsupported |
| `shape.list/create/delete/update` | Office.js ExcelApi 1.9：`list`；`create` 仅 `geometric`（rectangle\|ellipse\|triangle\|diamond\|rightArrow）或 `textBox`；`update` 浅层 newName/left/top/width/height/text/visible；text 仅 `hasText` 时回读。**无** image/line/group/fill/rotation。WPS → unsupported |

## 非目标

- 不替换 `desktop/` Electron 产品
- 不引入 COM / .NET / Electron 运行时
- 不伪造宿主不支持的能力（返回 typed `unsupported`）
- 不注册 workbook.open/create/save/switch、宏写入/运行、**UI 控件 / UserForm / 菜单**、Open XML 伪工具
- 本批不实现 dataValidation `errorAlert`/`showError`/`errorMessage`（字段拒绝）
