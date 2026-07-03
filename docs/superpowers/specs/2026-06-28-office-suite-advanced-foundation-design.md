# Office 三件套高级能力统一基础层设计

## 背景

当前项目已经具备 Excel COM 工具、Word/PPT COM 工具，以及 `.docx/.pptx/.xlsx` 的 Open XML 文件级检查、文本替换、布局检查、表格检查、表格样式和视觉快照选择外壳。下一步目标不是继续零散增加单点工具，而是建立统一基础能力层，让 Excel、Word、PPT 的高级复杂操作都遵循同一套 action、对象定位、结果、验证和监控协议。

用户确认第一阶段采用“底座 + 三件套各一批高级工具”的路线，首批能力包为“文档生产增强包”。

## 目标

第一阶段交付一个统一 Office 高级操作底座，并接入 Excel、Word、PPT 各一批高频文档生产能力：

- Excel：图表、条件格式、数据验证/下拉列表、表格样式增强。
- Word：标题样式批量应用、目录生成接口、表格美化、图片插入/替换、基础页眉页脚。
- PPT：主题/配色应用、幻灯片版式统一、图表插入/编辑、图片占位替换、形状对齐与分布。

这批能力必须能被模型自主选择和调用，并能被右侧编辑监控展示执行摘要、能力状态和验证结果。

## 非目标

- 不在第一阶段实现所有 Office 高级对象能力，例如宏工程编辑、SmartArt 深度编辑、动画时间轴完整编辑、母版完整编辑、数据透视表完整设计器。
- 不做万能兜底。Open XML 不能稳定处理的对象必须返回 `needsCom` 或 `unsupported`，不能假装成功。
- 不把 Excel、Word、PPT 三条线拆成互不兼容的工具协议。
- 不重构已有 Excel/Word/PPT COM bridge 的全部内部实现，只在必要处接入统一 action 协议。

## 核心设计

### 统一 Office Action 协议

新增统一 action schema，用于表达模型想对 Office 文件或当前 Office 应用执行的动作：

- `inspect`：读取结构、对象、样式、表格、图表、图片和占位符。
- `edit`：修改已有对象内容或属性。
- `style`：应用样式预设或局部样式。
- `insert`：插入图表、图片、表格、目录、页眉页脚等对象。
- `snapshot`：生成视觉快照，或返回不可用原因。
- `validate`：验证编辑结果，例如对象存在、数量变化、输出文件生成、目标样式写入。

每次 action 返回统一结果：

- `status`: `done | unsupported | needsCom | failed`
- `engine`: `openxml | com`
- `app`: `excel | word | presentation`
- `operation`: 具体操作名
- `filePath` / `outputPath`
- `target`
- `summary`
- `changes`
- `validation`
- `error`

### 对象定位

统一对象定位字符串用于跨三件套表达目标：

- Excel：`sheet:Sheet1`、`range:Sheet1!A1:D10`、`chart:1`、`table:1`
- Word：`heading:1`、`paragraph:12`、`table:1`、`image:logo`、`header:default`
- PPT：`slide:1`、`shape:Title 1`、`pictureSlot:1`、`chart:1`、`table:1`

第一阶段只要求 locator 可被工具解析和回显；高级定位可逐阶段增强。

### 能力声明

每个高级操作必须声明：

- 支持的 app 和文件类型。
- 首选 engine。
- 是否会写文件。
- 是否需要用户批准。
- Open XML 支持状态。
- COM 兜底条件。

模型提示词和 executor 不直接猜测能力，而是通过 registry/contract 暴露的工具描述和结果状态判断下一步。

## 架构边界

### `officeCore/`

新增统一基础层，职责是定义和校验跨三件套共享协议：

- action schema
- object locator
- result status
- validation result
- capability declaration

该层不直接读写文件，也不直接调用 COM。

### `officeOpenXml/`

继续负责文件级 `.docx/.pptx/.xlsx` 的 ZIP/XML 读写：

- Word/PPT/Excel 结构检查
- 表格和样式操作
- 图表、图片、主题、目录等可稳定文件级处理的对象
- 输出副本和 diff/validation 信息

Open XML 是首选路径，必须优先尝试文件级可测试实现。

### `officeCom/` 或现有 `office/`、`excel/`

负责 COM 兜底：

- 当前打开窗口操作
- Office 原生对象刷新
- Open XML 不稳定或不可覆盖的复杂对象

COM 不作为默认路径。只有用户明确要求操作当前打开窗口，或 Open XML 返回 `needsCom` 时才使用。

### `tools/registry`

暴露模型可见工具定义。第一阶段建议新增少量统一工具，而不是为每个对象无限扩展工具名：

- `office.action.inspect`
- `office.action.apply`
- `office.action.validate`

已存在的专用工具可以保留，并逐步接入统一结果协议。

### `tools/executors`

只做参数校验和路由，不承担业务逻辑。executor 根据 action 的 app、operation、target 和 preferEngine 调用 Open XML 或 COM adapter。

### `officeEditEvents`

扩展侧边栏监控摘要：

- 展示 action status。
- 展示 engine。
- 展示 changes。
- 展示 validation。
- 对 `unsupported / needsCom / failed` 给出清晰原因。

## 首批高级能力范围

### Excel

- `insertChart`：基于 range 创建基础图表，首批支持柱状图、折线图、饼图。
- `applyConditionalFormatting`：对 range 添加基础条件格式，例如大于/小于、色阶、数据条。
- `setDataValidation`：对 range 设置列表、数值范围或日期范围。
- `styleTable`：对 range/table 应用表格样式、表头样式、数字格式。

### Word

- `applyHeadingStyles`：按段落文本或层级规则应用标题样式。
- `insertOrUpdateToc`：插入目录字段或标记需要 COM 刷新。
- `styleTables`：美化表格表头、边框、隔行底色。
- `insertOrReplaceImage`：插入图片或替换指定图片占位。
- `setHeaderFooter`：设置基础页眉页脚文本。

### PPT

- `applyTheme`：应用配色和字体预设。
- `normalizeLayouts`：统一标题、正文、图片区域布局。
- `insertChart`：插入或更新基础图表数据和样式。
- `replacePictureSlot`：替换图片占位符。
- `alignShapes`：对齐和分布形状。

## 执行策略

1. 用户给出文件路径且不要求操作当前打开窗口时，优先走 Open XML。
2. 用户明确要求操作当前打开窗口，才走 COM。
3. Open XML 不支持当前对象时，返回 `needsCom`，由模型决定是否请求 COM 兜底。
4. 写操作默认输出副本，除非用户明确要求覆盖源文件。
5. 每个写操作必须附带 validation，至少验证输出文件存在和目标对象/样式有变化。
6. 每个阶段完成后运行聚焦测试、review、提交，再进入下一阶段。

## 错误处理

统一错误状态：

- `unsupported`：该对象或文件结构第一阶段不支持。
- `needsCom`：Open XML 能定位需求，但需要 Office 原生对象或当前窗口刷新。
- `failed`：执行异常，例如文件不存在、XML 损坏、参数非法。

禁止把 `unsupported` 自动包装成成功；禁止无条件调用 COM 作为隐藏兜底。

## 测试策略

- `officeCore`：纯单元测试，覆盖 action schema、locator 解析、结果状态和 validation。
- `officeOpenXml`：用最小生成的 `.xlsx/.docx/.pptx` zip 包做文件级测试。
- `executors`：mock adapter，验证路由和参数。
- `prompts`：验证提示词包含统一 action、Open XML 优先、COM 兜底。
- `officeEditEvents`：验证侧边栏摘要覆盖 `done / unsupported / needsCom / failed`。

## 分阶段建议

第一阶段拆成 6 个实现阶段：

1. 建立 `officeCore` 协议和测试。
2. 新增统一 action 工具注册和 executor 路由。
3. 接入 Open XML adapter，先复用现有 layout/table/style/snapshot 能力。
4. 接入 Excel 首批高级能力。
5. 接入 Word/PPT 首批高级能力。
6. 更新提示词、侧边栏监控和最终验证。

## 自检

- 需求覆盖：已覆盖统一基础层、Excel/Word/PPT 首批高级能力、Open XML 优先、COM 兜底、监控和验证。
- 范围控制：第一阶段不包含所有 Office 高级对象，不做万能兜底。
- 模块职责：`officeCore` 只定义协议，Open XML/COM adapter 执行具体能力，executor 只路由。
- 歧义处理：首批高级能力已明确为文档生产增强包，而不是完整 Office 自动化平台。
