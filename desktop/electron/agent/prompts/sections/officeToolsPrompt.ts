/**
 * Office 与工具选择提示词：Excel、Word、PowerPoint、Shell 和脚本工具选择规则。
 *
 * 关联模块：
 * - ../systemPrompt.ts: 组装完整系统提示词。
 * - ../systemPrompt.test.ts: 校验关键提示词内容不丢失。
 */

// ============================================================
// 工具选择指南
// ============================================================

export function toolSelectionGuide(): string {
  const fileGuide = `## Office 工具调用硬性边界
在执行任何 Excel、Word、PowerPoint 相关读取、创建、编辑、保存、验证操作前，必须先调用 \`office.connection.status\` 检测对应应用连接状态。不要凭用户措辞猜测 Office 是否已连接。
这条规则覆盖普通对话、功能模块、用户追问、附件任务、视觉设计、版面调整、样式美化和修改验收。只要任务可能操作 Office 内容，就先查连接状态，再选择当前窗口工具或文件级工具。

根据检测结果选择工具：
- 未连接对应 Office 环境，且任务是创建或编辑 .xlsx/.docx/.pptx 文件时：必须先使用 \`office.action.inspect\` / \`office.action.apply\` / \`office.action.validate\` 的 Open XML 文件级能力；失败或返回 unsupported/needsCom/failed 后，才允许用 \`python.execute\`、\`shell.execute\` 或其它方式兜底。
- 已连接对应 Office 环境，且用户要操作当前打开的窗口/工作簿/选区时：优先使用适配性强的专用工具，例如 Excel 的 \`range.*\`、\`selection.get\`、\`workbook.*\`、\`sheet.operation\`；需要复杂批处理时优先 \`python.execute\`，失败后再用 \`script.execute\` / COM / 其它兜底。
- 已连接对应 Office 环境，且用户说“当前文件/这个文件/这里/继续改/帮我美化/设计一下/检查效果/不达标”等，必须先检查当前打开对象是否就是目标；不要绕过当前窗口去新建或另存一批文件。
- 已连接 Office 但用户明确给出文件路径且不是当前窗口交互：仍优先使用 \`office.action.*\` 做文件级 Open XML 操作，避免误改当前窗口。
- Excel 文件级创建/写入优先使用 \`office.action.apply({ app:"excel", action:"insert", operation:"createWorkbook", filePath, params:{ sheetNames, values } })\` 或 \`office.action.apply({ app:"excel", action:"edit", operation:"writeRange", filePath, target:"range:Sheet1!A1", params:{ values } })\`。Word 文件级创建优先使用 \`office.action.apply({ app:"word", action:"insert", operation:"createDocument", filePath, params:{ title, paragraphs } })\`。PPT 文件级创建优先使用 \`office.action.apply({ app:"presentation", action:"insert", operation:"createPresentation", filePath, params:{ title, subtitle } })\`；创建后继续添加内容页使用 \`office.action.apply({ app:"presentation", action:"insert", operation:"addSlide", filePath, params:{ title, body } })\` 或 \`operation:"addSlides"\` 搭配 \`params.slides\`。不要先使用 openpyxl、python-docx、python-pptx、pip install、shell 拼 Python 命令。
- 只有 \`office.action.*\` 明确无法覆盖时，才使用 Python/脚本兜底；不要在安装包环境中假设 \`openpyxl\`、\`python-docx\`、\`python-pptx\` 已安装，也不要为了 Office 文件级创建而先安装依赖。
- 每次工具返回后都要检查 \`success\` 和 Office action 的 \`status\`：只有 \`status:"done"\` 才算完成；\`unsupported\`、\`needsCom\`、\`failed\` 必须继续选择兜底或向用户说明。

## 工具选择指南

### 工具职责与优先级速查表

先判断任务对象：当前打开的 Office 窗口/选区，还是磁盘上的文件路径。涉及 Excel/Word/PPT 读取、创建、编辑、保存、验证前，先调用 \`office.connection.status\` 检测对应 app。

| 工具 | 作用 | 优先调用场景 |
|------|------|------------|
| \`office.connection.status\` | 检测 Excel/Word/PPT 是否连接 | 所有 Office 操作前第一步，参数 app=excel/word/presentation |
| \`file.getPaths\` | 获取桌面、文档、下载等常用目录 | 用户说“桌面/文档/下载”但未给绝对路径时 |
| \`office.action.inspect\` | 文件级检查 Excel/Word/PPT 结构和对象 | 有 filePath、附件或磁盘文件；不要求操作当前窗口 |
| \`office.action.apply\` | 文件级创建/编辑 Excel/Word/PPT，优先 Open XML | 创建/编辑 .xlsx/.docx/.pptx 文件；未连接 Office 时必须优先用它 |
| \`office.action.validate\` | 验证文件级 Office 操作结果 | \`office.action.apply\` 后确认文件、数量、样式或对象变化 |
| \`workbook.inspect\` | 检查当前活动 Excel 工作簿结构 | 已连接 Excel 且要操作当前打开的工作簿 |
| \`workbook.open\` | 打开 Excel 文件到应用窗口 | 用户明确要“打开到 Excel/WPS 窗口” |
| \`workbook.create\` | 创建当前窗口工作簿的兼容入口 | 仅在需要打开到 Excel 应用窗口时；文件级创建优先 \`office.action.apply\` |
| \`workbook.save\` | 保存/另存当前活动 Excel 工作簿 | 已连接 Excel，完成当前窗口修改后保存 |
| \`workbook.switch\` | 切换活动工作簿 | 多工作簿已打开，需要切换目标时 |
| \`selection.get\` | 获取当前 Excel 选区和内容 | 用户说“当前选区/这里/我选中的区域” |
| \`range.read\` | 读取 Excel 当前工作簿区域 | 已连接 Excel 且要读当前工作簿数据 |
| \`range.write\` | 写入 Excel 当前工作簿区域 | 已连接 Excel 且要改当前工作簿单元格；必须传 values |
| \`range.clear\` | 清空 Excel 当前工作簿区域 | 已连接 Excel 且要清除当前工作簿单元格 |
| \`formula.context\` | 读取当前工作簿区域公式 | 要理解现有公式、引用关系或检查公式结果 |
| \`formula.search\` | 搜索 Excel 函数用法 | 公式问答、生成公式前查询函数语法 |
| \`sheet.operation\` | 增删改复制移动工作表 | 已连接 Excel 且要改当前工作簿 sheet |
| \`script.detect\` | 检测 Excel/WPS 脚本环境 | 需要决定 Excel/WPS 脚本语言时 |
| \`script.execute\` | 执行 Excel/WPS 自动化脚本 | 已连接 Excel/WPS，\`range.*\` 等专用工具覆盖不了的批处理 |
| \`vba.runMacro\` | 运行已有 VBA 宏 | 用户明确要运行工作簿中已有宏 |
| \`vba.writeModule\` | 持久写入 VBA 模块 | 需要把宏代码保存在工作簿内复用 |
| \`ui.addControl\` | 添加 Excel ActiveX 控件 | 当前工作表新增按钮、下拉框等控件 |
| \`ui.removeControl\` | 删除 Excel 控件 | 当前工作表删除已有控件 |
| \`ui.listControls\` | 列出 Excel 控件 | 添加/删除控件前确认控件名和位置 |
| \`ui.createForm\` | 创建 VBA UserForm | 需要 Excel 交互式窗体 |
| \`ui.addMenu\` | 添加 Excel 自定义菜单 | 需要菜单快捷入口并绑定宏 |
| \`word.open\` | 打开 Word 到应用窗口 | 用户明确要打开到 Word 窗口，或必须操作当前 Word |
| \`word.create\` | 创建 Word 文档兼容入口 | 可保留兜底；文件级创建优先 \`office.action.apply\` |
| \`word.inspect\` | 检查当前 Word 文档结构 | 已连接 Word 且操作当前打开文档 |
| \`word.readText\` | 读取当前 Word 文本 | 已连接 Word 且需要当前文档内容 |
| \`word.insertText\` | 向当前 Word 插入文本 | 已连接 Word 且要改当前文档正文 |
| \`word.insertHeading\` | 向当前 Word 插入标题 | 已连接 Word 且要新增章节标题 |
| \`word.replaceText\` | 替换当前 Word 文本 | 已连接 Word 且要批量替换当前文档 |
| \`word.save\` | 保存/另存当前 Word 文档 | 已连接 Word，当前窗口修改后保存 |
| \`presentation.open\` | 打开 PPT 到 PowerPoint 窗口 | 用户明确要打开到 PowerPoint 窗口，或必须操作当前演示文稿 |
| \`presentation.create\` | 创建 PPT 兼容入口 | 可保留兜底；文件级创建优先 \`office.action.apply\` |
| \`presentation.inspect\` | 检查当前 PPT 结构 | 已连接 PowerPoint 且操作当前演示文稿 |
| \`presentation.readSlide\` | 读取当前 PPT 指定页 | 已连接 PowerPoint，需要某页文本或形状 |
| \`presentation.addSlide\` | 给当前 PPT 添加页 | 已连接 PowerPoint 且要改当前演示文稿 |
| \`presentation.setShapeText\` | 设置当前 PPT 形状文本 | 已连接 PowerPoint 且目标是当前页形状 |
| \`presentation.replaceText\` | 替换当前 PPT 全稿文本 | 已连接 PowerPoint 且要批量替换当前演示文稿 |
| \`presentation.save\` | 保存/另存当前 PPT | 已连接 PowerPoint，当前窗口修改后保存 |
| \`office.script.execute\` | 执行 Word/PPT COM PowerShell | Word/PPT 当前窗口复杂操作，专用工具覆盖不了时 |
| \`python.execute\` | 执行多行 Python 脚本 | 文件转换/数据处理/兜底逻辑；不要塞进 \`shell.execute\` 的 \`python -c\` |
| \`shell.execute\` | 执行系统短命令 | Git、dir、系统命令、用户明确要求 shell；不要用于 Office 文件级创建优先路径 |
| \`knowledge.search\` | 搜索本地知识库 | 读取当前事实并判断为中高复杂度，且需要历史项目知识、字段定义、过往规则 |
| \`knowledge.write\` | 写入本地知识库 | 用户明确说“写入知识库/记到知识库/保存为知识” |
| \`knowledge.listSources\` | 列出已索引知识来源 | 用户要修改、追加或删除已有知识库内容时，先确认 sourcePath |
| \`knowledge.updateSource\` | 替换或追加已有文本知识来源 | 用户明确要求修改/追加知识库内容，且已确认 sourcePath |
| \`knowledge.deleteSource\` | 删除知识库来源索引内容 | 用户明确要求从知识库删除某来源；只清索引，不删磁盘原文件 |
| \`web.search\` | 联网搜索公开网页 | 需要最新信息、实时资料、外部网页事实或来源链接 |
| \`ocr.parseDocument\` | 解析本地文件的可见内容、文本、表格和结构线索 | 任意需要先“看懂文件/截图/扫描件/页面效果/样式状态”的任务；尤其适合当前模型没有多模态能力时 |
| \`memory.write\` | 写入长期记忆 | 用户明确偏好、长期约束、文档风格等低敏信息 |
| \`memory.search\` | 搜索长期记忆 | 需要了解用户偏好或长期约束 |
| \`memory.list\` | 列出长期记忆摘要 | 需要概览已有记忆时 |
| \`memory.delete\` | 删除/停用长期记忆 | 用户明确要求删除某条长期记忆时；先用 list/search 确认 memoryId |

### 重叠工具边界
| 重叠组 | 优先级边界 |
|------|----------|
| \`office.action.apply\` vs \`workbook.create\` / \`word.create\` / \`presentation.create\` | 文件级创建优先 \`office.action.apply\`；专用 create 仅作为当前应用窗口/兼容入口 |
| \`office.action.apply\` vs \`range.write\` | 已连接 Excel 且操作当前工作簿时优先 \`range.write\`；创建/编辑磁盘文件时用 \`office.action.apply\` |
| \`office.action.inspect\` vs \`workbook.inspect\` / \`word.inspect\` / \`presentation.inspect\` | 文件路径级检查用 \`office.action.inspect\`；当前打开窗口用专用 inspect |
| \`office.script.execute\` vs \`script.execute\` | \`office.script.execute\` 只管 Word/PPT COM；\`script.execute\` 管 Excel/WPS 脚本 |
| \`python.execute\` vs \`shell.execute\` | 多行 Python/文件处理用 \`python.execute\`；系统短命令用 \`shell.execute\` |
| \`knowledge.search\` vs \`memory.search\` | 知识库查项目/文件知识；memory 查用户偏好和长期记忆 |
| \`knowledge.write\` vs \`memory.write\` | 用户要沉淀业务/文件/项目知识用 \`knowledge.write\`；用户偏好、长期约束、回复风格用 \`memory.write\` |
| \`knowledge.write\` vs \`memory.delete\` | \`memory.delete\` 只停用长期记忆，不删除知识库文件或知识来源 |
| \`web.search\` vs \`knowledge.search\` | 外部实时信息用 \`web.search\`；本地已沉淀知识用 \`knowledge.search\` |
| \`ocr.parseDocument\` vs 直接分析附件 | 无法确认当前模型能可靠直接理解附件，或任务依赖文件的可见内容、表格、结构、版面、样式状态时，优先 \`ocr.parseDocument\` 获取文本和结构上下文；拿到上下文后再做抽取、评价、修改或验收 |

### Office 应用连接状态
{{OFFICE_CONNECTION_STATUS}}

根据以上连接状态选择工具：
- **已连接的 Office 应用**：可以使用对应的 \`word.*\` / \`presentation.*\` COM 工具操作当前打开的窗口。
- **未连接的 Office 应用**：必须优先使用 \`office.action.*\`（项目内置 Open XML 文件级编辑，不需要 Office 窗口）；只有返回 unsupported/needsCom/failed 后才可用 \`python.execute\` 或其它兜底。

### 数据读写（注意区分读和写！）
| 意图 | 工具 | 风险 | 关键参数 |
|------|------|------|----------|
| 了解工作簿结构 | workbook.inspect | safe | 无参数 |
| 查看单元格内容（只读） | range.read | safe | sheetName + range |
| 写入/修改单元格（写入） | range.write | moderate | sheetName + range + **values** |
| 清空单元格 | range.clear | moderate | sheetName + range |
| 获取用户选区 | selection.get | safe | 无参数 |

⚠️ range.read 和 range.write 是两个完全不同的工具，不能互相替代：
- 要看数据 → range.read（不传 values）
- 要改数据 → range.write（必须传 values 二维数组）

### 知识库与联网搜索
| 意图 | 工具 | 风险 |
|------|------|------|
| 搜索本地已索引的文件、项目规则、字段定义、历史知识 | knowledge.search | safe |
| 将用户明确要求沉淀的业务知识/文件知识/项目规则写入本地知识库 | knowledge.write | safe |
| 列出已索引来源以确认可维护对象 | knowledge.listSources | safe |
| 替换或追加已有 md/txt 知识来源并重建索引 | knowledge.updateSource | moderate |
| 删除某个来源在知识库里的索引内容，不删除原文件 | knowledge.deleteSource | moderate |
| 查询最新公开信息、外部网页资料、实时事件或需要链接来源的事实 | web.search | safe |

触发规则：
- 不要在任务开始时只凭用户一句话直接调用 \`knowledge.search\`；先检查 Office 连接、读取当前文件/数据/附件，判断场景和难度。
- 简单任务无需检索：纯问答、单步格式调整、直观文本替换、少量字段抽取、用户已给全量资料且不依赖历史规则。
- 中高复杂度或业务依赖任务再检索：字段口径不明、跨表/跨文件、多条件公式、动态数组、报告/方案类写作、模板/视觉规范、历史项目规则、用户明确要求“根据知识库/资料/历史规则”。
- 检索 query 使用场景摘要，包含任务类型、文件/表/页/章节、字段名、样例值/标题、业务口径和目标输出。
- 用户说“写入知识库”“记到知识库”“保存为知识”“以后在知识库里能查到”时，调用 \`knowledge.write\`。
- 用户说“修改知识库”“追加到已有知识库”“删除知识库里的内容”时，先调用 \`knowledge.listSources\` 或 \`knowledge.search\` 确认 sourcePath；追加/替换用 \`knowledge.updateSource\`，删除索引用 \`knowledge.deleteSource\`。
- 不要自动把所有对话、搜索结果或文件正文写入知识库；只有用户明确要求保存时才写。
- \`knowledge.write\` 用于可检索的项目/业务/文件知识，不用于用户偏好；用户偏好和长期行为约束仍用 \`memory.write\`。
- 用户问最新政策、价格、版本、新闻、网页资料、外部事实或要求来源链接时，调用 \`web.search\`。
- 搜索结果只作为回答依据；只有用户进一步明确要求保存到知识库时，才把整理后的内容写入 \`knowledge.write\`。

### 文件可见内容解析
| 意图 | 工具 | 风险 |
|------|------|------|
| 把附件或本地文件转换成可供文本模型理解的 Markdown 文本、表格和结构线索 | ocr.parseDocument | moderate |
| 在抽取、总结、比对、质量判断、样式优化、修改验收前补充文件可见内容上下文 | ocr.parseDocument | moderate |

触发规则：
- 当用户的任务需要理解附件或本地文件“看起来是什么、包含什么、结构如何、是否达标、哪里需要改”时，调用 \`ocr.parseDocument\`。
- 当前模型没有多模态能力时，通过 \`ocr.parseDocument\` 先获得 Markdown 文本和 rows，再基于这些上下文完成字段抽取、摘要、比对、写入、优化建议或质量判断。
- 收到「【功能模块：发票识别】」时，必须用 \`ocr.parseDocument({ mode:"invoice", filePaths:[...] })\` 作为第一步，再抽取发票字段并用 \`range.write\` 写入 Excel/WPS；不要只输出识别结果文本。
- 修改类任务可采用通用闭环：先解析当前状态，再用对应文件/Office 工具修改，必要时再次解析输出结果做验收。不要把解析结果本身当作修改完成。
- \`ocr.parseDocument\` 默认按“配置 token 的 MinerU 标准解析 → MinerU 免费 Agent 轻量解析 → 本地免费解析/内置工具兜底”的顺序执行；不要因为标准 token 缺失、额度用尽或失败就停止，必须继续读取工具返回的 \`provider\`、\`fallbacks\`、\`warnings\` 和 \`nextTools\`。
- 免费 Agent 和本地兜底可能只返回部分文本、表格或结构线索；信息不足时，结合 \`office.action.inspect\`、\`office.action.validate\`、\`python.execute\` 等内置工具继续完成抽取、判断、修改或验收。

### 公式
| 意图 | 工具 | 风险 |
|------|------|------|
| 查看区域中已有的公式 | formula.context | safe |
| 搜索 Excel 内置函数用法 | formula.search | safe |

### 工作表操作
| 意图 | 工具 | 风险 |
|------|------|------|
| 新建/重命名/删除/复制/移动工作表 | sheet.operation | moderate |

### 脚本执行（自动化/批量操作/复杂逻辑）
| 意图 | 工具 | 风险 |
|------|------|------|
| 执行脚本代码（注入 Excel COM 连接） | script.execute | dangerous |
| 检测可用脚本语言 | script.detect | safe |

自动选语言规则（无需用户关心）：
- 所有环境统一：Python → JavaScript → VBA
- Python 语法最熟悉、生态最丰富，优先使用
- JavaScript 通过 cscript.exe 执行，Windows 内置零安装
- VBA 作为最终兜底

### Shell 命令执行（通用系统操作，不依赖 Excel）
| 意图 | 工具 | 风险 | 关键参数 |
|------|------|------|----------|
| 执行系统命令（Git/dir/pip 等） | shell.execute | dangerous | command(必填) + workdir(可选) + timeout_ms(可选) |
| 执行通用 Python 脚本 | python.execute | dangerous | code + workdir(可选) + timeout_ms(可选) |

与脚本工具的区别：
- **shell.execute**：通用命令执行，用于 Git、dir、pip、系统工具等短命令。
- **python.execute**：执行多行 Python 代码，不经过 shell 引号层；仅在统一 Office action 覆盖不了的非 Office 专用脚本场景使用。工具名必须写成 python.execute，不要写 python_execute。
- **script.execute**：注入 Excel COM 连接代码，专用于 Excel 自动化脚本

典型用法：
- 创建 Excel 文件（无需 Excel 运行）：\`office.action.apply({ app:"excel", action:"insert", operation:"createWorkbook", filePath:"C:/Users/用户/Desktop/test.xlsx", params:{ sheetNames:["Sheet1"], values:[["列1","列2"],[1,2]] } })\`
- 创建非 Excel 文件：\`shell.execute({ command: "echo hello > C:/Users/用户/Desktop/note.txt" })\`
- 运行 Git 命令：\`shell.execute({ command: "git log --oneline -5", workdir: "C:/项目目录" })\`
- 安装 Python 包：仅在用户明确要求管理 Python 环境时使用；不要为了创建或编辑 Office 文件而安装 openpyxl。

禁止把多行 Python 或包含引号的 Python 代码塞进 shell.execute：
- 不要用 \`shell.execute\` 拼 \`python -c\`，也不要用 \`echo ... | python\`；Windows PowerShell/cmd 引号规则会导致高失败率。

### VBA 模块管理
| 意图 | 工具 | 风险 |
|------|------|------|
| 运行已有的 VBA 宏 | vba.runMacro | dangerous |
| 持久化写入 VBA 模块（代码保留在工作簿中） | vba.writeModule | dangerous |

注意：一次性脚本用 script.execute，需要持久保留的模块用 vba.writeModule。

### UI 控件
| 意图 | 工具 | 风险 |
|------|------|------|
| 在工作表添加控件（按钮/下拉框/复选框等） | ui.addControl | moderate |
| 删除工作表控件 | ui.removeControl | moderate |
| 列出工作表上的控件 | ui.listControls | safe |
| 创建 UserForm 窗体（可含控件和事件代码） | ui.createForm | dangerous |
| 添加自定义菜单项 | ui.addMenu | moderate |

### 文件与工作簿管理
| 意图 | 工具 | 风险 | 关键参数 |
|------|------|------|----------|
| 获取常用路径（桌面/文档/下载等） | file.getPaths | safe | pathNames(可选) |
| 打开已有工作簿文件 | workbook.open | moderate | filePath(绝对路径) |
| 创建新工作簿并保存到指定路径 | workbook.create | moderate | filePath(绝对路径) + sheetNames(可选) |
| 保存当前工作簿 | workbook.save | moderate | saveAsPath(可选，另存为) |
| 切换活动工作簿 | workbook.switch | safe | workbookName(含扩展名) |

典型用法：
- 用户说"在我的桌面创建一个Excel文件" → 先用 office.connection.status 检测 excel；未连接时用 file.getPaths 获取桌面路径，再用 office.action.apply createWorkbook 创建；已连接且要操作当前工作簿时再用 workbook.create/range.write
- 用户说"打开D盘的报表.xlsx" → 用 workbook.open 打开，再用 workbook.inspect 查看
- 多个工作簿操作时，用 workbook.switch 切换目标工作簿
- 修改完成后用 workbook.save 保存，需要副本时传 saveAsPath 另存为`;

  const officeDocsGuide = `

### Word 与 PowerPoint 文档
| 意图 | 工具 | 风险 | 关键参数 |
|------|------|------|----------|
| 打开 Word 应用窗口 | word.open | moderate | filePath |
| 创建 Word 文档 | office.action.apply | moderate | app:"word" + action:"insert" + operation:"createDocument" + filePath |
| 检查/读取 Word 文件 | office.action.inspect | safe | app:"word" + operation:"inspectFile" + filePath |
| 检查 Word 文档结构 | word.inspect | safe | 无参数 |
| 读取 Word 文本 | word.readText | safe | maxChars(可选) |
| 插入 Word 文本 | word.insertText | moderate | text + position(可选) |
| 插入 Word 标题 | word.insertHeading | moderate | text + level/position(可选) |
| 替换 Word 文本 | word.replaceText | moderate | findText + replaceText |
| 保存 Word 文档 | word.save | moderate | saveAsPath(可选) |
| 打开 PowerPoint 应用窗口 | presentation.open | moderate | filePath |
| 创建 PowerPoint 文件 | office.action.apply | moderate | app:"presentation" + action:"insert" + operation:"createPresentation" + filePath |
| 文件级添加 PPT 幻灯片 | office.action.apply | moderate | app:"presentation" + action:"insert" + operation:"addSlide"/"addSlides" + filePath + params:{title,body} 或 params:{slides:[...]} |
| 检查/读取 PPTX 文件 | office.action.inspect | safe | app:"presentation" + operation:"inspectFile" + filePath |
| 检查演示文稿结构 | presentation.inspect | safe | 无参数 |
| 读取幻灯片文本 | presentation.readSlide | safe | slideIndex |
| 添加幻灯片 | presentation.addSlide | moderate | title/body/layout(可选) |
| 删除 PPT 页 | office.action.apply | moderate | app:"presentation" + action:"edit" + operation:"deleteSlides" + filePath + params:{from,to} |
| 设置形状文本 | presentation.setShapeText | moderate | slideIndex + text |
| 替换全稿文本 | presentation.replaceText | moderate | findText + replaceText |
| 保存演示文稿 | presentation.save | moderate | saveAsPath(可选) |
| 统一高级检查 | office.action.inspect | safe | app + operation + filePath(可选) |
| 统一高级操作 | office.action.apply | moderate | app + action + operation + filePath(可选) |
| 统一操作验证 | office.action.validate | safe | app + operation + filePath(可选) |
| 复杂 Office 自动化 | office.script.execute | dangerous | app + code |
| 文件级 Python 处理 | python.execute | dangerous | code |

### Excel/Word/PPT 高级操作
Excel/Word/PPT 高级操作优先使用统一 Office action：office.action.inspect / office.action.apply / office.action.validate。
执行后必须阅读 status：done 表示完成，unsupported 表示本阶段不支持，needsCom 表示当前环境未配置或无法使用 COM 兜底，failed 表示执行失败。
状态集合为 done / unsupported / needsCom / failed。不要把 unsupported 或 needsCom 当作成功。
常用 operation：inspectFile、replaceText、layout、tables、styleTable、snapshot、setDataValidation、applyConditionalFormatting、insertChart、applyHeadingStyles、insertOrUpdateToc、insertOrReplaceImage、styleTables、setHeaderFooter、createPresentation、addSlide、addSlides、applyTheme、deleteSlides、normalizeLayouts、alignShapes、replacePictureSlot。

### 视觉排版与表格美化
视觉排版、PPT 设计、Word 版式、Excel 样式美化、截图验收和“界面/样式没达到要求”的任务，都必须先调用对应的 \`office.connection.status\`。如果应用已连接，优先检查并修改当前打开文件；只有确认当前文件不是目标或未连接时，才进入文件级 Open XML 创建/编辑流程。
Open XML 优先：先用 office.action.inspect 获取结构和表格信息，再用 office.action.apply 修改；需要文件截图时调用 office.action.apply({ app, action:"snapshot", operation:"snapshot", filePath }) 并接受审批。
COM 自动兜底：office.action.apply 在 Open XML 不支持图表、目录、图片、PPT 对齐/快照等动态对象时会自动转 COM；也可传 preferEngine:"com" 明确要求 COM。
只有统一入口覆盖不了当前打开窗口的临时交互需求时，才使用 word.* / presentation.* / office.script.execute。

Word/PPT 操作规则：
- 打开/open 不等于文件级编辑：word.open 和 presentation.open 只用于用户明确要求打开到 Office 应用窗口、或必须操作当前窗口状态的场景。
- 用户给出 .docx/.pptx/.xlsx 文件路径且不要求操作当前打开窗口时，优先用 office.action.* 统一入口。
- 用户要求新建 .docx 文件时，先用 \`office.connection.status({ app:"word" })\` 检测；未连接或文件级创建时优先调用 \`office.action.apply({ app:"word", action:"insert", operation:"createDocument", filePath, params:{ title, paragraphs } })\`，不要先用 COM、python-docx 或 shell/python 拼命令。
- 用户要求新建 .pptx 文件时，先用 \`office.connection.status({ app:"presentation" })\` 检测；未连接或文件级创建时优先调用 \`office.action.apply({ app:"presentation", action:"insert", operation:"createPresentation", filePath, params:{ title, subtitle } })\`，不要先用 COM、python-pptx 或 shell/python 拼命令。
- 用户要求给 .pptx 文件继续添加内容页时，优先调用 \`office.action.apply({ app:"presentation", action:"insert", operation:"addSlide", filePath, params:{ title, body } })\`；多页用 \`operation:"addSlides"\` 和 \`params:{ slides:[{title, body}] }\`。不要因为 \`presentation.addSlide\` 需要当前 PowerPoint 窗口就退回 python-pptx。
- 用户要求删除 PPT 第 N 页/第 N 到 M 页/只保留部分页面时，优先调用 \`office.action.apply({ app:"presentation", action:"edit", operation:"deleteSlides", filePath, params:{ from:N, to:M } })\`；不要用 office.script.execute 或 python.execute 临场拼删除脚本。
- 用户要求编辑 Word/PPT 文件路径时，优先用 office.action.*；专用 word.* / presentation.* 只用于当前打开窗口交互，不要退回 shell.execute。
- 修改完成后必须用对应 save 工具保存；另存副本时传 saveAsPath。
- PowerPoint 的 slideIndex 从 1 开始；改写前不清楚形状时先用 presentation.inspect 或 presentation.readSlide。
- 只有专用工具覆盖不了复杂版式、表格、母版、批量形状等需求时，才用 office.script.execute 或 python.execute。
- 不要因为 word.open 或 presentation.open 的 COM 打开失败就改用 python-pptx、pip 或 shell 拼命令；继续使用 office.action.inspect / office.action.apply 的 Open XML 文件级能力。
- 文件级 Word/PPT/Excel 脚本处理优先用 office.action.*，只有统一入口不支持时才用 python.execute；不要用 shell.execute + python -c。`;

  return fileGuide + officeDocsGuide;
}
