# 更新日志

这里只记录用户能够直接感知的功能与体验变化。代码重构、内部实现调整、测试数量和工程优化不进入更新日志。

## 未发布

### Excel 加载项

- 独立 Excel 加载项 Phase55.3：template print 严格拒绝 undefined/缺失成员（仅显式 null=unavailable）；defaultForAllPages.load 必填；freeze address 必须可解析 A1；工具数仍 98；**尚未真实 Excel/WPS 侧载验收**。
- 独立 Excel 加载项 Phase55.2：`workbook.template.capture` 固定少量 batch sync（非 O(表数)）、capture 专用严格 print snapshot；`apply` 禁止多单元格 bulk text、写后完整 header/freeze 回读、写前 surface 预检；Fake sync-gated + everReadBeforeSync；工具数仍 98；**尚未真实 Excel/WPS 侧载验收**。
- 独立 Excel 加载项 Phase55.1：`workbook.template.apply/capture` 合同收口（sheetNames minItems+空数组 fail-closed；plan/readback address+counts；严格 isNullObject；写前 surface 含 freeze getLocation；quote-aware A1；capture 单 Excel.run + 上下文内 pageLayout；坏 scalar fail-closed）；工具数仍 98；**尚未真实 Excel/WPS 侧载验收**。
- 独立 Excel 加载项 Phase55：`workbook.template.apply`（desktop `applyWorkbookTemplate` 对等；Office.js ExcelApi 1.8；写后严格回读；`autoFitVerified:false`）与 `workbook.template.capture`（desktop capture/inspect 浅层快照；ExcelApi 1.9；单工具无同义 inspect）；工具数 96→98；WPS typed unsupported；**尚未真实 Excel/WPS 侧载验收**。
- Phase54.2：selectedKeys/item key 唯一性与 isFilterCleared↔全选一致性；ClientResult 显式 lastClientResultReadBeforeSync 断言。
- Phase54.1：切片器宿主回读严格化（禁止 String/Number/Boolean 强制转换假绿）；filter apply/clear 验证全选与 isFilterCleared；list 不存在 sheet 失败；ClientResult 需 sync 后读 value。
- 新增 Office.js 切片器工具（`slicer.list` / `slicer.create` / `slicer.update` / `slicer.delete` / `slicer.filter.get|apply|clear`，ExcelApi 1.10 稳定子集）；创建需 `advancedIntent=interactive-pivot`，支持 Table/PivotTable 源；宿主无法回读来源字段时返回 `requestedSource` 与 limitations；`selectItems([])` 表示全选。WPS 仍为 typed unsupported。**尚未真实 Excel/WPS 侧载验收。**

- 独立 Excel 加载项 Phase53.3：DataValidation type/operator/alertStyle 宿主枚举回读改为仅大小写不敏感的精确匹配（enum exact readback）；**未**真实侧载验证。

- 独立 Excel 加载项 Phase53.2：DataValidation 完整 surface 类型校验 + 纯 planDvRuleWrite 写前计划；**未**真实侧载验证。

- 独立 Excel 加载项 Phase53.1：DataValidation errorAlert/prompt 按 Office.js ClientObject 合同预读 load+sync、整对象赋值与严格完整 snapshot；**未**真实侧载验证。

- 独立 Excel 加载项 Phase53：`dataValidation.read/write` 支持 Office.js `errorAlert`/`prompt` 与 `allowBlank`↔`ignoreBlanks` 完整回读；WPS 对新元数据 typed unsupported；**未**真实侧载验证。

- 独立 Excel 加载项 `chart.series.dataLabels` 合同收口：position 公开类型仅规范枚举；宿主 position 只接受官方 11 token（禁止模糊归一化）；1.8 路径写前预检全部成员；bubbleSizes 矩阵措辞与 dataLabels 对齐。
- 独立 Excel 加载项 `chart.series.dataLabels.update` 扩展 showPercentage/showBubbleSize/showLegendKey/separator/position（ExcelApi 1.8 `ChartSeries.dataLabels` 完整快照；position 不含 Invalid；separator 原样保留）；enabled-only 仍 ExcelApi 1.7；WPS 仍 unsupported；真实侧载尚未验收。
- 独立 Excel 加载项补强 `pivot.refresh` 版本回归：ExcelApi 1.3 与 1.7 门禁独立（1.7 关闭时 omit/false 仍可刷新透视）。
- 独立 Excel 加载项 `pivot.refresh` 支持 `refreshConnections=true`：调用 Office.js `Workbook.dataConnections.refreshAll`（ExcelApi 1.7；仅官方支持连接范围；`verified:false`；非完整 Workbook.RefreshAll）；false/省略仍仅 `PivotTable.refresh`（1.3）；WPS 仍 unsupported；真实侧载尚未验收。
- 独立 Excel 加载项 `chart.axes.update` 扩展 minorUnit（`""`=自动）、major/minorTickMark、tickLabelPosition（ExcelApi 1.7）、position/`setPositionAt`（ExcelApi 1.8）、linkNumberFormat（ExcelApi 1.9）；写后宿主回读；趋势线 label 仍暂缓；WPS 仍 unsupported；真实 Excel/WPS 侧载尚未验收。
- 修正 `chart.series.trendlines.format.update` 的 weight 合同：官方仅声明 points、无 min/max，执行层只校验有限数，不再无证据拒绝 ≤0。

- 独立 Excel 加载项新增 `chart.series.trendlines.format.update`（Office.js ExcelApi 1.7 `ChartTrendline.format.line` 颜色/线型/粗细写后回读；WPS 仍 unsupported；趋势线 label 格式未纳入；真实 Excel/WPS 侧载尚未验收）。

- 独立 Excel 加载项新增 `chart.series.markers.update`（Office.js ExcelApi 1.7 写后回读 series 标记样式/尺寸/颜色；WPS 仍 unsupported；真实 Excel/WPS 侧载尚未验收）。
- 硬化图表趋势线合同：intercept 支持 `""` 自动截距且回读为数字；add 使用返回对象 + getCount 确定 1-based 索引；文档明确 host getItem 为 0-based。

- 独立 Excel 加载项新增图表 series 趋势线工具 `chart.series.trendlines.list/add/update/delete`（Office.js ExcelApi 1.7/1.8 写后回读；WPS 仍 unsupported；series formula 仍无官方合同；真实 Excel/WPS 侧载尚未验收）。

- 独立 Excel 加载项 `chart.axes.update` 扩展坐标轴 displayUnit/customDisplayUnit、scaleType/logBase、showDisplayUnitLabel（ExcelApi 1.7）与主/次网格线可见性（ExcelApi 1.1）；写后宿主回读；series formula/categoryFormula 仍无 Office.js 合同故不实现；WPS 仍 unsupported；真实 Excel/WPS 侧载尚未验收。

- 独立 Excel 加载项透视表复审：执行层拒绝非 dataFields 的 function/caption；禁止零字段创建；refresh 门禁 ExcelApi 1.3、list/create 仍 1.8；`refreshConnections` 明确非桌面对等拒绝；默认 Pivots 落点对齐桌面 lastBottom+3；dataFields 允许多聚合同名字段；真实侧载尚未验收。
- 独立 Excel 加载项新增透视表生命周期：`pivot.list`/`pivot.create`/`pivot.refresh`（Office.js ExcelApi 1.8；空 destination→Pivots 表；写后 hierarchy 回读；`refreshConnections`/切片器 unsupported；WPS typed unsupported；真实 Excel 侧载尚未验收）。
- 独立 Excel 加载项页面布局补齐桌面打印对等：`printArea`/`printTitleRows|Columns`（别名 `repeatRows|Columns`）写后回读；`fitToOnePageWide|Tall` 映射为 fitToPages=1；打印区域/标题 **clear** 仍无官方 Office.js 合同故拒绝空串；WPS pageLayout 仍 unsupported；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项新增 `workbook.save`：保存当前已打开工作簿（Office.js ExcelApi 1.1；WPS `ActiveWorkbook.Save` 成员探测）；无路径/saveAs/打开/创建/切换；真实 Excel/WPS 侧载尚未验收。
- 开发脚本 `npm run dev:http` 改为跨平台（`npm_lifecycle_event` / `VITE_DEV_HTTP`），修复 Windows cmd.exe 无法解析 `VITE_DEV_HTTP=1 vite` 的问题。
- 独立 Excel 加载项交付门禁：区分提示词/文档中的 desktop·Electron 文本说明与运行时 import/require；WPS/生产包扫描禁止 child_process/Electron/COM/.NET 运行时依赖；WPS 剩余能力审计见 docs（无凭猜测扩实现）。

- 文档纠偏：Office.js `chart.source.update` 同表/跨表 A1 已实现（非 cross-sheet unsupported）；WPS JSA 开发指引默认 `https://localhost:3000`。

- 独立 Excel 加载项 WPS 数据验证：between/notBetween 缺 Formula2 时 read 诚实返回 unsupported（rule=null）；已有规则时 Delete 失败不再继续 Add；真实侧载尚未验收。

- 独立 Excel 加载项 WPS JSA 条件格式/数据验证：**implemented***（FormatConditions 1-based index id、Validation 写前快照与 Add 失败恢复；复用 classifyListSource/dvRulesMatch；成员缺失 typed unsupported）；真实 WPS 侧载尚未验收。

- 独立 Excel 加载项修复数据验证单值内联列表：host source `Yes`/`1`/`x` 分类为 inline 而非 unsupported；clear 测试强制验证 setup 写入成功；恢复 8 个 CF 运算符与 date/time DV 及 custom 成功 round-trip；真实 Excel 侧载尚未验收。

- 独立 Excel 加载项条件格式/数据验证：list 源仅无损同簿 A1 才标记 range；命名范围/函数/外部引用为 unsupported；custom 拒绝宿主额外 formula2；真实侧载尚未验收。
- 独立 Excel 加载项进一步收紧条件格式/数据验证：list Range 必须 load 地址、公式比较最小归一化（owner 表上下文）、list 源仅同簿 A1、非 between 拒绝宿主 formula2；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项加固条件格式/数据验证宿主回读：CF NotEqualTo 官方 token、add 校验规则与颜色、DV 全规则与 allowBlank 回读、list Range 代理真实地址、clear 要求 hostType=None；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项收紧条件格式与数据验证：CF 列表诚实返回 hostType（不把 ContainsText/DataBar 等伪装为 cellValue），add 仅 cellValue/custom 且完整比较运算符；DV 补齐 list/wholeNumber/decimal/date/time/textLength/custom，list 区域源传 Range 代理，Inconsistent/MixedCriteria 标记 limitations；ExcelApi 1.6/1.8 预检；写后宿主回读；WPS CF/DV 后续已按 FormatConditions/Validation 落地为 implemented*（见上条）；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项新增只读工具 `workbook.objects.inspect`：一次聚合 sheets/tables/charts/namedRanges/shapes 分类清单（可截断并保留 totalCount）；单分类失败/不支持不拖垮整项；WPS 上 table/chart/shape 仍明确 unsupported；真实 Excel/WPS 侧载尚未验收。
- 独立 Excel 加载项接入公式治理工具：依赖检查、引用修复、公式转值、备份检查/恢复（WENGGE_FORMULA_BACKUP_V1 隐藏表；文本方式保存公式原文；真实 Excel/WPS 侧载尚未验收）。
- 独立 Excel 加载项新增公式保护检查与锁定/解锁（仅公式单元格 + 可选表保护）；密码仅请求内存使用且不进入工具结果；WPS 明确不支持；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项表格现支持首末列高亮（showFirstColumn/showLastColumn），以及表格 AutoFilter 筛选与多级排序工具；WPS 仍明确返回不支持，真实 Excel 侧载尚未验收。
- 独立 Excel 加载项可生成正式本地 WPS JSA jsaddons 包（相对路径任务窗格 + publish/ribbon/entry）；真实 WPS 侧载尚未验收。
- 独立 Excel 加载项现在会恢复模型供应商、Gateway 和活动供应商配置；API Key 仍仅保存在当前任务窗格内存中，不会写入浏览器持久化存储。
- 独立 Excel 加载项支持同源 AI Gateway 连接模式：可配置 gatewayBaseUrl 与 gatewayUpstreamId，浏览器不保存或不发送 API Key，请求经同域 `/api/ai/v1/:upstreamId/...` 转发。
- 独立 Excel 加载项的表格更新现可调整同一工作表内的表格范围，并切换行/列交替底纹；操作完成后会回读宿主中的实际表格范围与显示状态。
- 独立 Excel 加载项新增区域行列结构操作：可插入、删除并自动调整选区行高/列宽；自动调整完成后会回读宿主中的实际尺寸；WPS 对 insert/delete/autofit 为成员探测实现（implemented*），真实 Excel/WPS 侧载尚未验收。
- 独立 Excel 加载项图表数据源支持同表与跨工作表 Range（含带空格/引号表名），写后回读 series；仍拒绝外部工作簿、3D、多区域与结构化引用；真实 Excel 侧载尚未验收。
- 独立 Excel 加载项 WPS 路径新增工作表可见性、工作表保护/取消保护、命名区域增删改查与区域插入/删除（成员探测 + 单测；密码仅请求内存且不回显；真实 WPS 侧载尚未验收）。
- 新增独立同源 AI Gateway 服务（`ai-gateway/`）：浏览器可通过同域 `/api/ai/v1/:upstreamId/...` 访问服务端显式配置的上游，无需在客户端暴露 API Key。

## 0.1.86 - 2026-07-18

- 完整权限模式现在会自动执行所有已注册工具，不再为危险、删除、联网或未知工具弹出确认窗口；参数校验、路径授权和执行层硬拒绝仍然生效。
- 删除“远程数据处理”设置及其运行时开关；OCR、联网搜索、发票抽取和知识 Embedding 按实际能力直接运行，高置信凭据仍在外发前阻止。
- 更正权限模式文案为“完整权限（自动执行）”，与实际执行行为保持一致。
- 公式任务现在会优先采用“单个锚点公式 + spill”填充同一行级逻辑的列区域，不再默认生成逐格公式并要求用户下拉；仅在用户明确要求、宿主不支持或验证失败时降级。
- 同时打开 Microsoft Excel 与 WPS 表格时，当前窗口操作不再按固定顺序随机绑定宿主；未明确选择时会返回宿主歧义并要求选择，公式写入错误会显示实际绑定的宿主和 COM 原因。
- 公式写入不再依赖 Formula2；普通、现代和动态公式统一走 Formula，仅明确要求传统 CSE 数组公式时才走 FormulaArray。
- WPS 当前窗口的 `range.write` 恢复使用逐单元格默认 `Value` 入口，避免动态数组公式被统一 COM 公式属性改写。

## 0.1.85 - 2026-07-16

- Office 文件级工具现在直接向模型声明每个 operation 的准确 `params` 格式，包括必填/可选字段、枚举、数组和嵌套对象；不再依赖冗长场景提示或让模型猜造参数，同时模型工具上下文保持在约 53k，避免重新膨胀到 400k 以上。

## 0.1.84 - 2026-07-16

- 修复 WPS 文字和演示已打开有效文档时连接状态不稳定、WPS 演示被误判为 Microsoft PowerPoint 且无法识别有效进程的问题；状态探测现在复用活动文档枚举，并返回规范化的 WPS 宿主、实例和进程信息。

## 0.1.83 - 2026-07-16

- 修复创建新的 Excel、Word 或 PowerPoint 文件时错误依赖已打开桌面应用，或对不存在文件直接执行后续编辑的问题；独立磁盘文件现在先通过 Open XML 创建，多页 PPT 再批量添加页面。
- 修复 Office 工具定义重复展开导致新会话上下文占用超过 400k 的问题；模型可见工具预算降至约 15k，128k/256k 等上下文模型可正常使用，同时保留执行前的严格参数校验。

## 0.1.82 - 2026-07-16

- 本机会话、记忆与知识等恢复数据改为本地加密保存；设置中可轮换本地数据密钥。
- 更换数据目录后会尝试清理旧目录中的应用数据；清理失败则登记待擦除；导出与旧目录会纳入受管副本登记。
- 擦除本地数据会覆盖已登记的活动目录、旧目录与应用创建的隐私导出副本，并销毁本地数据密钥，返回可审计删除摘要；不删除原始 Office 文档、附件或知识源。
- 修复 Excel/WPS 写入现代数组公式时被 `@` 隐式交叉降级的问题，并可回读写入分类。
- 修复 Excel/WPS 多单元格传统数组公式只写入首个单元格的问题；写入或验证失败时会恢复完整目标区域。
- Excel 公式写入失败时会自动恢复目标区域，避免出现工具报错但部分单元格已经被改写。
- 图表插入现在会校验可见性、系列、位置、尺寸和单元格锚点；校验失败不再误报成功。
- 修复数据透视表创建与默认目标位置，未指定位置时使用专用 `Pivots` 工作表；Power Query、透视表和切片器的边界由提示词升级为执行层校验，普通写值/公式任务不再升级为复杂工具。
- Provider 密钥、OCR Token 与远程压缩 API Key 改为系统加密存储，设置界面只显示掩码；外链、危险操作和文件访问采用更严格的安全校验。
- 热补丁增加防重放、有效期和逐文件完整性校验，检测到已安装文件被修改时不再加载补丁。
- 新增默认关闭的“远程数据处理”总开关；关闭时 OCR 仅本地解析、联网搜索禁用、知识库自动降级为关键词检索，并在发送前阻止高置信密钥内容。
- 联网搜索会流式限制第三方响应体，异常超过 2 MiB 时安全拒绝，避免搜索服务异常响应持续占用内存。
- OCR 工具现在会准确标记单个文档的文本或表格是否被裁剪，避免把不完整结果误报为完整内容。
- 加强外部内容与长期记忆隔离：网页、OCR 和工具结果不再被当作指令，记忆写入必须引用当前轮用户明确表达的原文。
- 数据目录更换改为完整 staging、哈希校验和失败回滚，设置、会话、知识库、Office 备份/事务/工作流会一并迁移；默认数据改存用户隔离目录。
- 热补丁改为流式解压，并增加 Renderer 健康确认、白屏/崩溃回滚和签名吊销策略。
- 模型工具参数现在会在请求用户确认前和实际执行前统一校验；缺少必填项、未知字段、错误枚举或超出数量限制的调用会直接返回明确错误，不再进入审批或 Office Worker。
- 模型流已经显示正文、推理或工具状态后若网络中断，不再整体重试并重复显示已有内容；尚未产生可见输出的瞬时失败仍会自动重试。
- IPC 入口现在限制聊天长度、附件/OCR 文件数量、Excel 写入矩阵规模、设置值、模板变量和 Base64 文件传输大小；高成本通道按渲染进程限流，超限请求会在进入 Office Worker、读取大文件或写盘前被拒绝。
- 日志与工具执行审计不再保存参数、结果或密钥原文，历史检索索引也不再复制工具载荷和完整推理内容，降低本地敏感数据暴露面。
- Power Query、数据透视表和切片器改为按当前任务意图动态开放；普通写值、公式、格式和固定汇总请求的模型工具列表中不再出现这些高级 operation。
- 修复 WPS 文件级表格操作可能重复打开同一文件并最终超时的问题；WPS 不支持持久化 Power Query 时会返回明确错误，透视表和切片器保持可用。
- 优化桌面端首次打开速度；设置、Office 自动化和任务功能面板改为使用时加载。
- 本地日志、Office 备份、事务和工作流增加自动留存与容量上限；运行中记录会被保护，过期历史按固定周期清理，避免应用数据目录无限增长。
- 常用 Excel、Word 和 PowerPoint 文件级操作现在按应用与操作校验参数；Excel 图表、打印设置、公式治理、PDF 导出、工作簿预设和对象管理也会逐层校验，Excel/Word 同名导出操作不会串用参数；未知对象类型、字段、枚举、缺少替换规则和不安全的工作流变量名会在确认与执行前被拒绝。
- Word 引用、修订和文档比较现在只接受实际可执行的命令与参数；内容控件修订必须明确 tag 或 title，避免空选择器意外改写全部控件。
- Word 邮件合并和内容控件工具现在会拒绝 CSV 数据源、无效输出格式、虚构字段映射及未实现命令；删除和更新内容控件必须使用明确且不会歧义的选择器。
- Word 长文档排版现在必须明确选择自动标题识别或提供非空匹配条件，避免默认把全部段落改成标题；自定义标题正则增加长度和执行超时保护。
- PowerPoint 动画现在必须逐项明确目标形状、类型和效果，放映与讲义必须显式选择模式；未知效果、空批量备注和把输出路径误放进 params 的调用会提前拒绝。
- PowerPoint 品牌操作现在必须明确是否显示页码，布局操作必须选择策略和目标形状；空参数不会再默认开启页码或把整页内容套用网格排版。
- Excel 联动 Word/PPT 的导出、报告包、链接刷新和重绑现在按目标应用严格校验；增量更新必须提供稳定链接标识，报告 section 缺范围、图表名或重链接来源时会在执行前明确拒绝。
- Office 工作流变量和步骤结果占位符现在只读取真实自有属性，并限制路径段数与格式；原型链、保留字段、空路径段和无效步骤选择器会在执行前拒绝。
- 常规设置新增本地数据导出：可把会话、记忆、知识索引、Office 备份/自动化和日志校验复制到空目录，API Key、OCR Token 与自定义请求头秘密不会进入导出包。
- 常规设置新增受保护的应用本地数据擦除：输入精确确认短语后，可清除当前数据目录中的功能设置、会话、记忆、知识索引、Office 备份/自动化、日志和临时文件；数据目录位置会保留，也不会删除目录外的原始文档。

## 0.1.81 - 2026-07-12

- 修复热补丁下载时提示 `Invalid URL` 并导致无法安装的问题。
- 优化软件更新按钮的视觉样式与检查状态反馈。

## 0.1.80 - 2026-07-12

- 修复深色模式下已发送消息文字不可见的问题。

## 0.1.79 - 2026-07-12

- 新增软件更新中心，可在应用内检查、下载并覆盖安装新版本。
- 新增轻量热补丁更新，补丁下载完成后重启应用即可生效。
- 搜索窗口右上角增加关闭按钮，退出搜索更直接。
- 优化自动压缩阈值滑条，并将聊天输入上限提升至 50000 字符。
- 新增独立产品下载页和功能更新日志页面。

## 0.1.78 - 2026-07-12

- 新增 WPS、Excel、Word 和 PowerPoint 快速启动入口。
- 支持直接创建并写入 Excel/WPS 内部 VBA 或 JSA 宏。
- 优化窗口透明度、紧凑模式和侧边栏交互。
- 精简 Office 工具选择，降低重复调用和无效回退。

## 0.1.77 - 2026-07-11

- 优化公式方法论与环境识别，减少函数兼容性的错误判断。
- 动态数组结果支持自动读取完整溢出区域。
- 统一功能侧栏开关，并在公式需求提交后清空旧草稿。

## 更早版本

- 增加公式助手、代码生成、OCR、数据清洗、报告生成和图表制作模块。
- 支持 Excel、WPS、Word 与 PowerPoint 的当前窗口及文件级处理。
- 增加会话管理、上下文压缩、知识库、长期记忆和安全策略设置。
