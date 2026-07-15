# Office 高级自动化

## 调用边界

- 生产链路固定为模型工具 -> TypeScript 类型化薄桥 -> .NET 8 Worker -> COM/WPS JSA 或 C# Open XML，不允许绕过 Worker 拼接外部脚本。
- 当前 Excel/WPS 窗口或选区使用 `workbook.*`、`range.*`、`formula.*`、`sheet.*`。
- 当前 Word/PPT 窗口使用 `word.*`、`presentation.*`。
- 磁盘 `.xlsx/.docx/.pptx` 文件使用 `office.action.inspect/apply/validate`。Open XML 可完成时不启动 Office；动态对象、导出和应用对象模型使用 COM/WPS 兜底。
- 同一应用打开多个文件时，先调用 `office.documents.list` 取得完整路径，再用 `office.objects.list` 获取工作表、页面、幻灯片或对象的稳定 locator；激活时原样传入完整路径与 locator。
- 同时安装 Microsoft Office 与 WPS 时，文件级 COM 操作用 `params.host` 明确宿主：表格传 `excel/wps`，文字传 `word/wps`，演示传 `powerpoint/wps`。跨应用报告可分别传 `sourceHost`、`wordHost`、`presentationHost`。目标文件未打开时使用隔离进程，不附着无关活动窗口。

### Excel 基础与高级工具边界

- 直接写值、公式、格式或固定汇总结果始终优先 `range.read/write`；数据量本身不触发 Power Query、透视表或切片器。
- Power Query 仅用于外部/多来源的可刷新 ETL、联接与加载生命周期，并要求明确源、转换、加载位置和文件路径。
- 透视表仅用于用户明确要求的透视对象或交互式多维字段布局；固定分组汇总能用公式产出时仍走 `range.write`。
- 切片器只用于已有透视表或结构化表的交互筛选。
- 执行层不只依赖提示词：Power Query 必须传 `params.advancedIntent:"refreshable-etl"`，创建/更新时还要传 `sourceKind:"external"|"multi-source"`；透视表和切片器必须传 `params.advancedIntent:"interactive-pivot"`。缺少这些显式语义标记时工具在进入 Worker 前拒绝。
- 模型可见的文件级调用按 `app + operation` 校验参数。基础检查/验证、快照、Excel 图表插入与深度编辑/条件格式/数据验证/表格样式、Word 标题/目录/表格/页眉页脚/图片及 PPT 常用编辑已禁止未知 `params`；`formatChart` 的 `series`、`axes`、`dataLabels` 也逐层拒绝未知字段。尚未建模的其他 COM 深度操作保留兼容分支并继续受统一深度、节点、集合和字节预算限制。
- 工作流模板变量最多 128 个顶层键；键名仅允许字母或下划线开头，后续使用字母、数字、下划线或连字符。顶层键不得包含点号，嵌套值使用 `{{vars.customer.name}}` 引用。

## 高级 operation

### Excel/WPS 表格

| operation | 用途 | 关键 params |
| --- | --- | --- |
| `createPivotTable` | 创建数据透视表 | `rowFields`、`columnFields`、`filterFields`、`dataFields`、`destination`、`name` |
| `refreshPivotTables` | 刷新透视表及可选连接 | `refreshConnections` |
| `addSlicer` | 给透视表添加切片器 | `pivotName`、`field`、`name`、`caption`、位置参数 |
| `inspectPowerQueries` | 读取查询公式、说明、连接和工作表/数据模型加载状态 | 可选 `name` |
| `createPowerQuery` / `managePowerQuery` | 管理查询完整生命周期和加载目标 | `name`、`command`、`mFormula`、`loadMode`、`destination` |
| `inspectCharts` / `formatChart` | 检查并深度编辑图表、系列、坐标轴和标签 | `chartName/index`、`series`、`axes`、`dataLabels`、区域和尺寸参数 |
| `inspectWorkbookObjects` / `manageWorkbookObject` | 枚举并管理工作簿多类型对象 | `types`；或 `objectType`、`command` 及对象参数 |
| `manageWorksheetObjects` | 旧版形状管理兼容入口 | `command`、`name` 及位置尺寸属性 |
| `captureWorkbookTemplate` / `inspectWorkbookFormatting` | 捕获工作簿格式为可复用模板 JSON | `headerRows`、`titleRows`、`totalRows` |
| `applyWorkbookTemplate` | 应用内置预设或捕获模板及分层格式规则 | `preset`/`template`、`sheetNames`、工作表规则 |
| `inspectPrintSettings` / `configurePrint` | 回读或设置完整页面与打印参数 | `sheetNames`、纸张/方向/页边距、标题行列、缩放、分页符、页眉页脚 |
| `exportSheetsToPdf` | 合并或分别批量导出指定工作表 PDF | `sheetNames`、`mode`、`outputPath`/`outputDirectory`、`overwrite` |
| `traceFormulaDependencies` / `inspectFormulaDependencies` | 读取同表、跨表、外部、名称和结构化引用，构建正反依赖并检测循环、`#REF!` | `scope`、`target`、`maxExpandedRangeCells`；只读 inspect 操作 |
| `repairFormulaReferences` | 批量修复错误引用 | 明确 `replacements`，或 `copyFromNeighbors:true` 使用相邻公式 R1C1 模式 |
| `convertFormulasToValues` / `inspectFormulaBackups` / `restoreFormulas` | 备份后固化公式、查看备份并按批次恢复 | `target`/`scope`、`createBackup`、`backupId`、`removeAfterRestore` |
| `inspectFormulaProtection` / `manageFormulaProtection` | 检查、锁定或解锁公式区域并保护工作表 | `command:lock/unlock`、`password`、`unlockInputs`、`protectSheet` |
| `exportPdf` / `snapshot` | 导出工作簿/工作表 PDF 或区域 PNG | `scope`，以及 `target`/`outputPath` |

#### Power Query 数据管道

`managePowerQuery.command` 支持 `create`、`update`、`upsert`、`duplicate`、`rename`、`load`、`refresh`、`unload`、`delete`。`loadMode` 支持 `worksheet`、`dataModel`、`connectionOnly`；工作表加载使用 `destination:"Sheet2!A1"`，可用 `tableName` 指定输出表名。重命名会重建已有加载关系，卸载/删除可用 `clearOutput:true` 清理输出数据。

#### 图表深度编辑

`formatChart` 支持更换数据源、类型、标题、图例、样式和位置尺寸；`series` 支持新增、更新、删除、公式或数据范围、组合图类型、主次坐标轴、平滑线和数据标签；`axes` 支持类别/数值轴、主次轴、标题、最小/最大值、主单位、数字格式和逆序。趋势线、误差线、轴交叉值、对数轴、网格线以及图表区/绘图区的独立格式尚未实现，不应作为可用参数发送。

`insertChart` 在保存前回读图表对象、工作表可见性、系列数、位置尺寸和 `TopLeftCell/BottomRightCell`。任一关键检查失败会返回 `chart_verification_failed`，不会以成功 summary 掩盖。模型随后应使用 `office.action.validate` + `inspectCharts` 再次确认文件中的图表对象。

#### 数据透视表创建与验证

`createPivotTable` 未传 `destination` 时创建或复用专用 `Pivots` 工作表，并把新透视表放到已有透视表下方，避免覆盖源数据。创建后会回读缓存、行/列/筛选/数据字段及 `TableRange1/2`；缓存、字段或目标范围缺失时返回 `pivot_verification_failed`。文件级验收使用 `inspectWorkbookObjects` 并筛选 `pivotTable`。

#### 工作簿对象

`inspectWorkbookObjects` 可按 `types` 筛选工作表、名称、结构化表、图表、形状、连接、查询、透视表和切片器。`manageWorkbookObject.objectType` 支持 `worksheet`、`name`、`table`、`connection`、`shape`、`chart`、`image`、`pivotTable`、`slicer`，各类型提供创建、重命名、移动、缩放、显隐、刷新、样式或删除等适用命令。

#### 专业格式与模板

`captureWorkbookTemplate` 返回版本化模板，包含主题、工作表层级、表头样式和逐列格式，可直接作为其他工作簿的 `params.template`。`applyWorkbookTemplate` 内置 `professional`、`financial`、`dashboard`、`minimal` 预设，并支持工作表级 `titleRows/headerRows/totalRows`、主题、列规则、区域规则、结构化表样式、公式/数值/数据条/色阶/图标集条件格式、冻结窗格、网格线和打印规则。

#### 打印与 PDF

`configurePrint` 可同时处理多个 `sheetNames`。纸张支持 A3/A4/A5/B4/B5/Letter/Legal/Tabloid；`marginUnit` 支持 `centimeters`、`inches`、`points`。使用 `repeatRows`/`repeatColumns` 设置打印标题，`horizontalPageBreaks`/`verticalPageBreaks` 设置手动分页；`fitToOnePageWide:true` 固定一页宽，默认不强制一页高。`exportSheetsToPdf.mode` 为 `combined` 时按工作表顺序生成一个 PDF，为 `separate` 时输出到 `outputDirectory`。

#### 公式治理

`range.write` 对 Excel/WPS 共用同一公式分类与写入策略：普通公式走 `Formula`，现代/动态公式走 `Formula2`，只有显式 `legacyCse:true` 才走 `FormulaArray`。现代公式缺少 `Formula2` 时明确失败，不回退 `Formula`，避免重新引入 `@`。返回值包含 `written`、`dynamicCells`、`arrayCells`、`plainCells`。

文件级 Open XML 写入使用同一函数分类和前缀规范化；动态锚点写入数组公式引用与 `c@cm`，工作簿通过 `CellMetadataPart` 保存 `XLDAPR` 动态数组 metadata。多个动态锚点复用同一 metadata 描述，不写工作表 `extLst`，也不使用 `aca=false` 冒充动态数组能力。

依赖检查会返回 `nodes`、`edges`、`cycles[{path}]`、`brokenReferences` 和 Excel 当前循环引用位置。大范围引用超过 `maxExpandedRangeCells` 时保留为范围边，避免无界展开。

错误引用不会自动猜测。`repairFormulaReferences.replacements` 提供 `find`、`replace`，可选 `sheetName`；只有同一行或列的相邻公式模式可靠时才使用 `copyFromNeighbors:true`，并按 R1C1 相对结构复制。

`convertFormulasToValues` 默认在工作簿内创建 VeryHidden `_WenggeFormulaBackup*` 备份表并返回 `backupId`。普通公式和动态数组/旧数组公式按实际结果固化；`restoreFormulas` 可按 `backupId` 恢复公式、数字格式和锁定状态。该备份用于公式恢复，仍与文件级事务备份并存。

### Word/WPS 文字

| operation | 用途 | 关键 params |
| --- | --- | --- |
| `inspectDocumentFormatting` / `formatLongDocument` | 检查并统一九级标题、正文、引用、题注、表格和页面结构 | `headingRules/styles`、`normalStyle`、`quotePatterns`、`margins`、`sectionBreaks`、`headerFooter`、`toc` |
| `inspectReferences` / `manageReferences` | 检查并管理书签、脚注、尾注、题注、交叉引用和图表目录 | `command`、`name/text/label/item`、`targetType/index/position` |
| `inspectRevisions` / `manageRevisions` | 读取批注修订，按作者、类型、时间、范围或文本规则接受/拒绝 | `command`、`rule`、`enabled` |
| `compareDocuments` | 对比原文与修订文档，输出带修订的比较文档和结构化摘要 | `revisedFilePath`、`author`、`granularity`、`outputPath` |
| `applyTrackedChanges` | 在修订模式下执行 AI 替换、插入、删除和字段修改 | `edits`、`keepTracking` |
| `prepareMailMergeTemplate` | 把模板占位符转换为 MERGEFIELD 域 | `fields[{placeholder,field}]` |
| `mailMerge` / `batchMailMerge` | 使用 Excel/CSV 数据生成单个或逐记录文档 | `dataSourcePath`、记录范围、输出格式/目录、命名字段、条件和图片字段 |
| `inspectContentControls` / `populateContentControls` | 检查并按控件类型填充智能模板 | `values`、`fieldMap`、格式/列表/图片选项 |
| `manageContentControls` | 创建、删除和管理控件、列表项及锁定状态 | `command`、`controls`/`selector`、控件类型参数 |
| `exportPdf` | 导出 PDF | `outputPath` 可选 |

#### 样式与长文档排版

`formatLongDocument` 可按 `headingRules[{pattern,level}]` 或内置中英文编号模式自动设置 1-9 级标题；正文、引用和“图/表/公式”题注分别套用对应样式。`normalStyle`、`headingStyles` 支持字体、字号、粗斜体、间距、缩进、行距和段落保持；`clearDirectFormatting:true` 可清除正文直接格式后重新套用样式。

页面结构通过 `margins`、`orientation`、`sectionBreaks[{position,type}]`、`headerFooter` 和 `pageNumbers` 控制；`toc:create/update` 创建或刷新目录。操作结束会更新正文、页眉页脚、目录和图表目录中的域。

#### 书签与交叉引用

`manageReferences.command` 支持 `addFootnote`、`addEndnote`、`addBookmark`、`deleteBookmark`、`addCaption`、`addCrossReference`、`addTableOfFigures`、`updateFields`、`updateAll`。题注目标支持 `table`、`image/figure`、`equation` 或指定范围；交叉引用支持章节、编号项、书签、脚注、尾注及自定义“图/表/公式”标签。

#### 审阅与修订

`inspectRevisions` 返回每条修订和批注的作者、类型、日期、文本及范围。`manageRevisions.command` 支持 `track`、`acceptAll/rejectAll`、`acceptMatching/rejectMatching`、`deleteComments/deleteCommentsMatching`；`rule` 可按 `authors`、`types`、`fromDate/toDate`、`start/end`、`textPattern` 组合筛选。

`applyTrackedChanges` 会先开启 Word 修订模式，再执行 `replace/delete/insert/replaceBookmark/replaceContentControl` 编辑，因此 AI 修改保留可接受或拒绝的原文轨迹。`compareDocuments` 不覆盖原文和修订稿，而是使用段落级 LCS 生成独立差异报告，并返回新增、删除段落及样例摘要；超大文档自动退化为按段落位置比较，避免 Word 原生比较弹窗阻塞自动化。

#### 邮件合并与批量文档

`prepareMailMergeTemplate` 将 `{{姓名}}` 等占位符转换成 Word 邮件合并域。`batchMailMerge` 逐记录生成文件，`outputFormat` 支持 `docx`、`pdf`、`both`；`fileNameField` 或 `fileNamePattern:"{编号}-{姓名}"` 控制文件名，并自动替换 Windows 非法字符。

`conditions` 使用 `{placeholder,field,operator,value,trueText,falseText}` 在单条合并结果中处理条件内容；`imageFields` 使用 `{placeholder,field,width,height}` 将数据源中的图片路径写入文档。已有 Word IF 域也会由原生邮件合并处理。

#### 内容控件与智能模板

`inspectContentControls` 返回 ID、Tag、Title、类型、文本、选中状态、列表项和锁定状态。填充时优先按 Tag，再按 Title 和 ID 匹配 `values`；支持文本/富文本、复选框、日期、组合框、下拉列表和图片，默认保留控件原有字体格式。`fieldMap` 可将模板字段映射到业务数据键。

`manageContentControls.command` 支持 `add`、`delete`、`setLock`、`addListEntry`、`clearListEntries`、`setValue`。创建时可一次传入多个 `controls`，设置类型、Tag、Title、占位文本、列表项、默认值和锁定状态。

### PowerPoint/WPS 演示

演示 COM 操作可在 `params.host` 中传 `powerpoint` 或 `wps` 明确选择宿主；同时安装两套软件时应按用户当前使用的软件传值。未指定时按 PowerPoint、WPS 演示的可用顺序选择。

| operation | 用途 | 关键 params |
| --- | --- | --- |
| `inspectPresentationTheme` | 读取设计、母版、版式、主题色、页脚和页面尺寸 | 无 |
| `inspectSlideElements` | 读取文本框、图片、形状、表格、图表及坐标，并检测文字溢出、遮挡和越界 | `allSlides` |
| `insertTable` | 创建表格并一次写入二维数据 | `name`、`values`、`rows/columns`、`left/top/width/height` |
| `applyMasterBranding` | 应用模板并统一母版、版式、字体、Logo、页脚、页码和主题色 | `templatePath`、`fontName/fontMap`、`logoPath`、`footerText`、`themeColors`、`layoutMap` |
| `layoutElements` | 精确编辑、网格排版、对齐、等距分布、保持比例、裁剪和越界修复 | `edits`、`mode`、`shapeNames`、`align`、`distribute`、`fitToSlide` |
| `inspectAnimations` | 读取动画顺序、类型、触发、时长、延迟和页面切换 | `allSlides` |
| `configureAnimations` | 添加进入、强调、退出和路径动画 | `effects[].category/effect/shapeNames/trigger/order/duration/delay/pathX/pathY` |
| `configureSlideShow` | 配置自动播放、循环放映、放映类型、页面切换和翻页时间 | `showType`、`autoPlay`、`loop`、`transition`、`advanceAfter` |
| `setSpeakerNotes` | 写入单页或批量演讲者备注 | `text` 或 `notesBySlide[]` |
| `inspectSpeakerNotes` | 读取讲稿并检查缺失页和页面内容对应度 | `allSlides` |
| `exportHandouts` | 导出备注页或每页 1/2/3/4/6/9 张的 PDF 讲义 | `includeNotes`、`layout`、`outputPath` |

### 跨应用

| operation | 用途 | 关键 params |
| --- | --- | --- |
| `exportRangeToWord` | 将 Excel 区域或图表写入 Word，可保留数据源链接 | `linked`、`sourceType`、`chartName`、`linkId`、`title`、`overwrite` |
| `exportRangeToPresentation` | 将 Excel 区域或图表写入 PPT，可保留数据源链接 | `linked`、`sourceType`、`chartName`、`linkId`、`title`、`overwrite` |
| `buildReportPackage` | 从多个 Excel 区域/图表同时生成 Word 报告和 PPT 汇报 | `sections`、`linked`、`outputDirectory`、`baseName`、`overwrite` |
| `inspectLinkedOfficeContent` | 在 Word/PPT 中列出链接对象、来源和 locator | 无；只读检查 |
| `refreshLinkedOfficeContent` | 原位刷新 Word/PPT 中的 Excel 链接对象 | 无；不删除或重建页面对象 |

跨应用输出先写入同目录临时文件，全部成功后再发布到正式路径。覆盖已有输出时会临时保留旧文件，发布失败会恢复旧版本。

`linked:true` 使用链接 OLE。图表传 `sourceType:"chart"` 和 `chartName`；区域使用 `target:"range:工作表!A1:D20"`。刷新前工具只读打开链接源 Excel 并完成计算，再调用 Word/PPT 的 `LinkFormat.Update()`；链接对象在原位置更新，因此人工调整的页面、尺寸和排版不会被重建覆盖。

## 事务与工作流

原地修改前，适配器将源文件复制到应用数据目录下的 `office-backups`。结果中的 `data.transaction` 包含 `sourcePath`、`backupPath`、操作名和创建时间。

备份默认保留 30 天，每个源文件最多 50 份、全局最多 500 份且总计不超过 2 GiB。创建操作会保护本次新备份，之后的周期维护从最旧记录开始执行 TTL 和配额清理。备份元数据中的路径必须仍指向受控备份目录，伪造的外部路径不会被跟随或删除。

应用启动时及之后每 6 小时执行本地数据生命周期维护：应用日志默认保留 30 天、30 个文件、100 MiB；Office 事务默认保留 30 天、200 个目录、2 GiB；工作流默认保留 90 天、500 个记录、100 MiB。当前日期日志、租约仍有效的 `running` 工作流以及 90 天恢复宽限期内的 `pending` 事务受保护；`conflicted` 事务与 `paused` 工作流不会被配额提前挤出，但到期后仍可清理。损坏记录保留给人工恢复，不猜测删除。清理失败只记录告警，不阻止应用启动或 Office 操作。

- `office.action.inspect` + `listBackups`：列出全部或指定源文件的备份。
- `office.action.apply` + `restoreBackup`：通过 `params.backupPath` 恢复原文件。
- `office.workflow.run`：顺序执行最多 20 个文件级步骤，中间步骤可以产生明确文件；每步持久化状态、结果、产物和时间。
- 持久化工作流失败时默认进入 `paused` 并返回 `workflowId`、`failedStep`、`nextStep`；传 `resume:true` 和原 `workflowId` 从失败步骤继续，已成功步骤不会重跑。显式 `failureMode:"rollback"` 才会立即撤销。
- 每个工作流同时创建组事务并返回 `transactionId`。开始前快照全部已知受影响文件；撤销时恢复原文件并删除本次新建的声明产物，重做时按原步骤重新执行。
- `office.workflow.status` 查看步骤和失败位置；`office.transaction.list/inspect` 查看事务、产物和修改清单；`office.transaction.undo/redo` 整体撤销或重做。
- 执行后才发现、且任务开始前未声明的未知产物只记录在清单中，不会在撤销时冒险删除。

示例：

```json
{
  "steps": [
    {
      "app": "excel",
      "action": "style",
      "operation": "applyWorkbookTemplate",
      "filePath": "D:\\报表.xlsx",
      "target": "range:明细!A1:H200"
    },
    {
      "app": "excel",
      "action": "edit",
      "operation": "exportPdf",
      "filePath": "D:\\报表.xlsx",
      "outputPath": "D:\\报表.pdf"
    }
  ]
}
```

## 兼容性

- 文件级基础编辑优先使用 .NET Worker 内的 C# `DocumentFormat.OpenXml`，不要求 Office 进程运行；旧 TypeScript Open XML 实现已移除。
- 高级对象操作需要本机安装对应 Microsoft Office 或 WPS，并提供兼容 COM 对象模型。
- COM 清理按进程归属执行：本次新建的 Office 进程会完整退出；复用用户已有 WPS 进程时只关闭本次打开的文件并释放 COM 对象，不遍历关闭其他文件，也不调用应用级 `Quit()`。
- WPS 12.0 的主题色 COM 属性只读，且多页备注可能只持久化第一页。工具会在 WPS 保存并释放文件后使用 XML 解析器更新主题包或补齐备注部件及关系，再通过后续检查回读真实结果。
- Power Query、切片器、动画和讲义等能力在不同 Office/WPS 版本中的对象模型覆盖不同；不支持时工具会返回 `failed`，不会伪报成功。
- 工具不暴露任意 Shell、Python、PowerShell 或 JScript 执行入口；Office 自动化统一由类型化的 .NET Worker 协议执行。

安装了 Microsoft Office 的开发机可运行 `npm run test:office-smoke`、`npm run test:word-smoke` 和 `npm run test:presentation-smoke`，分别验证 Excel 深度能力、Word 排版/引用/修订/邮件合并/内容控件，以及 PPT 母版品牌、元素诊断、四类动画、放映、备注和讲义导出。`npm run test:office-reliability` 额外实测 Excel 链接 Word/PPT、原位刷新、流水线暂停续跑、跨文件撤销重做及多窗口对象选择。`npm run test:excel-lifecycle` 与 `npm run test:word-lifecycle` 验证同一文件连续打开和锁释放。

这些脚本只按变更范围定向执行，禁止把全套真实 Office 冒烟作为默认门禁长时间串行运行。生产 action 默认超时 120 秒；冒烟默认单动作 30 秒，并每 10 秒输出等待探测。可在 PowerShell 中设置 `$env:WENGGE_OFFICE_SMOKE_TIMEOUT_MS="45000"` 临时调整单动作时限。设置 `$env:PRESENTATION_SMOKE_HOST="wps"` 或 `"powerpoint"` 可明确演示宿主。

WPS 冒烟必须先在磁盘创建测试文件，再用 WPS 打开该现有文件；不得从 WPS 新建界面创建测试文件。运行前后只清理确认属于测试的残留进程，避免连接到错误实例或关闭用户正在编辑的文件。
