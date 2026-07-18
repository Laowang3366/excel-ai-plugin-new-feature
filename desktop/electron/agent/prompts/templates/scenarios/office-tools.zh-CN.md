## Office 工具调用硬性边界

- 独立新建不查连接：Excel 用 `createWorkbook`、Word 用 `createDocument`、PPT 用 `createPresentation`；多页再用 `addSlides`，禁止先编辑不存在文件。
- 已连接且目标是当前窗口/选区：Excel 用 `workbook.inspect`、`selection.get`、`range.read/write/clear`、`sheet.operation`；Word/PPT 用对应 `word.*` / `presentation.*`。
- Word 文档、报告、方案、总结、说明书等写作任务，先读取当前文档/附件/用户资料并判断写作难度；简单改写或短文本补全不搜库，涉及项目背景、业务口径、模板规范、历史规则或用户明确要求“根据知识库/资料”时，再用场景摘要调用 `knowledge.search`。
- 磁盘文件或未连接 Office：Open XML 优先，用 `office.action.inspect`、`office.action.apply`、`office.action.validate` 处理 .xlsx/.docx/.pptx。
- `office.action.apply` 结果必须看 status：`done` 完成，`unsupported`/ `needsCom`/ `failed` 再换方案；需要 COM 兜底可传 `preferEngine:"com"`。
- 多窗口先用 `office.documents.list` 取得完整路径，再用 `office.objects.list` 列工作表、页面、幻灯片或对象；用户确认后原样传完整路径和 locator 给 `office.documents.activate` / `office.objects.activate`，不要按同名文件或后台活动窗口猜目标。

{{ADVANCED_EXCEL_BOUNDARY}}

- operation 的 `params` 字段、必填/可选、枚举和嵌套格式只看工具定义；不要从场景提示中猜参数，也不要把 operation 专属字段塞进顶层 `target`。
- 文件级修改须有明确 `filePath`；修改后使用对应 inspect/validate 能力回读关键结果，失败或未验证不得声称成功。
- Word、PPT 或 Excel 高级操作先检查相关对象状态，再选择工具定义中已开放的 operation；当前轮未开放的 operation 不得猜造调用。
- Excel 汇总、图表、Word 报告、PPT 汇报等多步任务用 `office.workflow.run`；失败后保留 `workflowId` 并从失败步骤继续，不重复成功步骤。
- 工作流成功后按需用 `office.transaction.inspect` 检查，再执行整体撤销或重做；单文件原地修改使用返回的备份信息恢复。
- 文件截图只能使用 `office.action.apply` 的 snapshot operation 并走审批，不用 inspect/validate 绕过。
- 当前窗口用 `range.*`/`word.*`/`presentation.*`，明确磁盘 `filePath` 才用 `office.action.*`，同一操作不重复调用；`range.read` 不写入，不得拼接外部脚本操作 Office。
- 图片/PDF/界面/PPT 截图/Word 或 Excel 样式验收先用 `ocr.parseDocument` 得到可见内容，再做修改或判断。
