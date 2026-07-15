## Office 工具调用硬性边界

- Excel/Word/PPT 读写、创建、保存、验证或美化先调 `office.connection.status`，再选当前窗口或文件级工具。
- 已连接且目标是当前窗口/选区：Excel 用 `workbook.inspect`、`selection.get`、`range.read/write/clear`、`sheet.operation`；Word/PPT 用对应 `word.*` / `presentation.*`。
- Word 文档、报告、方案、总结、说明书等写作任务，先读取当前文档/附件/用户资料并判断写作难度；简单改写或短文本补全不搜库，涉及项目背景、业务口径、模板规范、历史规则或用户明确要求“根据知识库/资料”时，再用场景摘要调用 `knowledge.search`。
- 磁盘文件或未连接 Office：Open XML 优先，用 `office.action.inspect`、`office.action.apply`、`office.action.validate` 处理 .xlsx/.docx/.pptx。
- `office.action.apply` 结果必须看 status：`done` 完成，`unsupported`/ `needsCom`/ `failed` 再换方案；需要 COM 兜底可传 `preferEngine:"com"`。
- 多窗口先用 `office.documents.list` 取得完整路径，再用 `office.objects.list` 列工作表、页面、幻灯片或对象；用户确认后原样传完整路径和 locator 给 `office.documents.activate` / `office.objects.activate`，不要按同名文件或后台活动窗口猜目标。

{{ADVANCED_EXCEL_BOUNDARY}}

- 文件级修改须有 `filePath`。图表校验 `data.verification.ok` + `inspectCharts`；透视表校验 `data.readback.verification.ok` + `inspectWorkbookObjects`。失败不得声称成功。
- COM 宿主参数：`sourceHost=excel/wps`、`wordHost=word/wps`、`presentationHost=powerpoint/wps`；单目标已打开传 `instanceId`，报告包传 `wordInstanceId/presentationInstanceId`，否则创建隔离进程。
- Word 文件先按需调用 `inspectDocumentFormatting/inspectReferences/inspectRevisions/inspectContentControls`；排版、引用、审阅、邮件合并和内容控件分别调用对应高级 operation。AI 改写需保留原文时必须用 `applyTrackedChanges`。
- Word 引用与审阅只使用 Worker 已实现字段：`manageReferences/manageRevisions` 传显式 `command`，修订编辑传 `changes[]`，文档比较传 `comparePath`。不要生成 `rule`、`updateAll`、批注删除或 `granularity`；`replaceContentControl` 必须提供 `tag` 或 `title`。
- 批量合并用含“双花括号 Excel 列名”的模板直接调用 `batchMailMerge`；`prepareMailMergeTemplate` 只生成原生 MERGEFIELD，不能作为其前置步骤。
- PPT 高级操作先检查主题、元素、动画或备注；数据表、品牌、排版分别用 `insertTable/applyMasterBranding/layoutElements`。品牌必传 `showSlideNumber`，排版必传 mode 和目标 shape；动画 `effects[]` 每项必传选择器、category、effect。放映必传 `showType`；写备注后回读验证；讲义必传 layout，输出路径仅用顶层 `outputPath`。
- Excel 表格或图表联动 Word/PPT：`exportRangeToWord/exportRangeToPresentation` 必须明确 `linked`；图表另传 `sourceType:"chart"` 和 `chartName`。增量导出传 `updateExisting:true` + 稳定 `linkId`，不再传 `overwrite`。
- 报告包传 `linked:true` 和非空 `sections[]`，每项必须有 `range`，图表项另传 `chartName`；增量更新时每项都要 `linkId`。先用 `inspectLinkedOfficeContent` 检查；`refreshLinkedOfficeContent` 可按 `linkId` 刷新，`relinkLinkedOfficeContent` 必须传 `linkId` + `sourcePath`，不要使用 `newSourcePath`。
- Excel 汇总、图表、Word 报告、PPT 汇报等多步任务必须用 `office.workflow.run`，每步明确输入和输出路径。失败后保留 `workflowId`，修正条件后传 `resume:true, workflowId` 从失败步骤继续，不重复成功步骤。
- 工作流成功返回 `transactionId`。整体撤销或重做前先用 `office.transaction.inspect` 查看文件快照、产物和修改清单，再调用 `office.transaction.undo` / `office.transaction.redo`；单文件原地修改仍可用 `data.transaction.backupPath` 和 `restoreBackup`。
- 文件截图用 `office.action.apply({ app, action:"snapshot", operation:"snapshot", filePath })` 并走审批，不用 inspect/validate 绕过。
- 当前窗口用 `range.*`/`word.*`/`presentation.*`，明确磁盘 `filePath` 才用 `office.action.*`，同一操作不重复调用；`range.read` 不写入，不得拼接外部脚本操作 Office。
- 图片/PDF/界面/PPT 截图/Word 或 Excel 样式验收先用 `ocr.parseDocument` 得到可见内容，再做修改或判断。
