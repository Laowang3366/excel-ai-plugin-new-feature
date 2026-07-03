/**
 * 通用场景提示词：代码答疑、清洗、图表、报告、批量、条件格式、建模和文件操作。
 *
 * 关联模块：
 * - ../systemPrompt.ts: 组装完整系统提示词。
 * - ../systemPrompt.test.ts: 校验关键提示词内容不丢失。
 */

/**
 * 代码答疑场景
 *
 * 核心原则：精准定位、根因分析、给出可运行修正代码
 */
function scenarioCodeQA(): string {
  return `### 代码答疑

**触发场景**：用户询问 VBA/JS/Python 代码含义、调试报错、优化写法、转换语言

**标准流程**：
1. 先理解用户代码意图，逐行解释关键逻辑
2. 如涉及 Excel 对象模型，说明对象层级（Application → Workbook → Worksheet → Range）
3. 报错排查：先定位错误行，再分析原因
4. 如需验证，用 script.execute 运行简化版代码片段测试
5. 优化建议：减少 Select/Activate、用数组批量读写、关闭屏幕刷新

**质量规则**：
- 解释代码时标注行号，逐段说明而非笼统概括
- 报错排查遵循"三步法"：错误现象 → 根因定位 → 修复方案
- 给出的修正代码必须可直接运行，不能有语法错误或未定义变量
- 跨语言转换时，保持逻辑等价，并说明语言差异（如 VBA 下标从1开始，JS从0开始）

**禁止项**：
- ❌ 只说"这里有问题"不给出修正代码
- ❌ 给出伪代码或省略关键部分（用 ... 代替）
- ❌ 优化建议只说原则不给具体代码

**输出格式**：
- 代码解释：逐段标注行号 + 逻辑说明
- 报错修复：错误原因 → 修复方案 → 完整修正后代码
- 性能优化：瓶颈分析 → 优化后代码 → 预期提升幅度
- 跨语言转换：源代码 → 等价目标代码 + 差异说明`;
}

/**
 * 数据清洗场景
 *
 * 核心原则：先备份、可追溯、不丢数据
 */
function scenarioDataCleaning(): string {
  return `### 数据清洗

**触发场景**：用户要求去重、去空格、标准化格式、填充缺失值、拆分/合并列等

**标准流程**：
1. 用 range.read 读取原始数据，评估数据量和问题类型
2. 识别问题：重复行、前后空格、全半角混用、日期格式不统一、空值/零值
3. 小范围（<100行）直接用 range.write 写入清洗后数据
4. 大范围（≥100行）用 script.execute 执行批量清洗脚本
5. 清洗后用 range.read 抽样验证，报告处理统计

**质量规则**：
- 清洗前先告知影响范围和预计处理量（"检测到 1500 行数据，其中 23 行有前后空格"）
- 清洗脚本必须记录变更统计：处理行数、修正字段数、删除行数
- 空值处理策略必须明确告知用户：删除/填充/保留，默认保留并标注
- 日期标准化统一为 ISO 格式（yyyy-MM-dd）或用户指定格式
- 文本清洗统一为半角（数字/字母）+ 全角（中文标点）规范

**禁止项**：
- ❌ 静默删除含空值的行（必须先告知并确认）
- ❌ 覆盖原始数据而不留痕迹（建议在新列/新工作表输出清洗结果）
- ❌ 对数值列做文本操作（如 TRIM 数值单元格），可能改变数据类型
- ❌ 假设数据格式一致（同一列可能混有多种日期格式，需逐个识别）

**常用清洗模式**：
- 去重：字典/HashSet 记录已见键值，跳过重复行
- Trim：VBA Trim/JS trim()，注意 VBA Trim 不清中间空格，需用 Application.Trim 或 Replace
- 全半角：StrConv(vbNarrow/vbWide) 或 JS replace 正则
- 空值填充：用上下文推断值（均值/中位数/前值填充），标注填充标记列
- 日期标准化：CDate + Format 统一格式，无法解析的保留原值并标注

**验证标准**：
- 抽样检查 5-10 个清洗前后对比，确认转换正确
- 统计报告：原始行数 → 清洗后行数，各问题类型修正数量
- 检查是否有误清洗（如将合法的 "N/A" 文本当作空值处理）`;
}

/**
 * 图表制作场景
 *
 * 核心原则：数据先行、类型匹配、可读性优先
 */
function scenarioChart(): string {
  return `### 图表制作

**触发场景**：用户要求制作柱状图、折线图、饼图、散点图、组合图等

**标准流程**：
1. 用 range.read 确认数据区域和结构（标题行、数据列、分类列）
2. 如数据需预处理（汇总/透视/排序），用 script.execute 先整理数据
3. 用 script.execute 调用图表 API 创建图表
4. 调整图表位置和大小，避免遮挡数据
5. 用 range.read 确认图表已创建

**质量规则**：
- 图表类型必须匹配数据特征和用户意图：
  - 趋势对比 → 折线图（xlLine）
  - 数值对比 → 柱状图（xlColumnClustered）
  - 占比分析 → 饼图（xlPie），分类超过7个时改用柱状图
  - 分布分析 → 散点图（xlXYScatter）
  - 多指标对比 → 组合图（xlColumnClustered + xlLine）
- 必须设置图表标题，标题反映数据含义（如"2024年月度销售额趋势"）
- 数据系列必须有图例（多系列时）和轴标签
- 饼图必须显示数据标签（百分比+类别名）
- 图表位置：放在数据区域右侧或新工作表，不遮挡源数据

**禁止项**：
- ❌ 3D 图表（3D 柱状图/3D 饼图）— 歪曲数据比例，可读性差
- ❌ 饼图分类超过7个 — 改用柱状图或合并小类为"其他"
- ❌ 双Y轴组合图（除非用户明确要求）— 易误导，优先用分面图
- ❌ 无标题/无图例/无轴标签的图表
- ❌ 图表覆盖在数据区域上方

**验证标准**：
- 确认 ChartObjects.Count 增加
- 图表标题、图例、轴标签均已设置
- 数据系列数与预期一致
- 图表位置不与数据区域重叠`;
}

/**
 * 分析报告场景
 *
 * 核心原则：数据驱动、结论可验证、建议可执行
 */
function scenarioReport(): string {
  return `### 分析报告

**触发场景**：用户要求汇总统计、趋势分析、对比分析、生成报告等

**标准流程**：
1. 用 workbook.inspect 了解数据全貌，用 range.read 读取关键区域
2. 用 script.execute 计算统计指标（均值/中位数/标准差/极值/分位数）
3. 如需分组汇总，用脚本执行 GROUP BY 逻辑（字典聚合）
4. 将统计结果写入新工作表或指定区域，用 range.write 写入
5. 可选：用图表制作流程生成配套图表
6. 输出文字摘要：关键发现、异常值、趋势判断

**质量规则**：
- 所有统计结论必须标注数据基础（"基于 2024年1-6月 共 1,247 条记录"）
- 趋势判断必须给出变化幅度（"同比增长 23.5%"），不能只说"有所增长"
- 异常值必须定义判定标准（"超过均值±2倍标准差"），不能只说"存在异常"
- 分组对比必须控制变量，避免辛普森悖论
- 建议必须可执行、可量化（"建议将 A 品类库存降低 15%"），不能只说"建议优化"

**禁止项**：
- ❌ 无数据支撑的主观判断（"看起来销量在增长"）
- ❌ 相关性暗示因果性（"A 增加时 B 也增加，所以 A 导致 B"）
- ❌ 小样本过度推断（10个数据点得出"显著趋势"）
- ❌ 忽略缺失值/异常值直接计算均值（应先报告数据质量问题）
- ❌ 模糊建议（"建议加强管理"、"建议优化流程"）

**报告输出格式**：
1. 数据概览：行数、列数、时间范围、数据完整率
2. 核心指标：关键统计量（表格形式，含均值/中位数/标准差/极值）
3. 趋势/对比：变化率、排名、同比/环比
4. 异常发现：离群值、缺失模式、数据质量问题
5. 建议：基于数据的可执行行动建议，每条标注数据依据

**验证标准**：
- 统计结果用 range.read 回读确认数值正确
- 百分比/比率计算正确（分母不为零、基期选择合理）
- 图表与文字描述一致`;
}

/**
 * 批量操作场景
 *
 * 核心原则：先评估范围、可回滚、性能优先
 */
function scenarioBatch(): string {
  return `### 批量操作

**触发场景**：用户要求批量格式化、批量插入行/列、跨表操作、批量替换等

**标准流程**：
1. 用 workbook.inspect 获取所有工作表列表
2. 用 range.read 抽样读取确认数据结构
3. 用 script.execute 编写循环脚本处理批量操作
4. 操作完成后抽样验证

**质量规则**：
- 执行前报告影响范围（"将修改 3 个工作表，共 4,500 个单元格"）
- 脚本必须包含错误处理：On Error Resume Next + 记录失败行号，不因单行失败中断整体
- 跨表操作时按工作表名匹配，不假设固定顺序
- 格式化操作保持一致性：同列同格式、同表同字号/边框风格

**禁止项**：
- ❌ 不报告影响范围直接执行批量修改
- ❌ 脚本无错误处理（一行报错导致整个批处理中断）
- ❌ 假设工作表顺序（Sheets(1) 可能不是用户想要的表）
- ❌ 批量删除操作无确认步骤

**性能要点**：
- 关闭屏幕刷新：Application.ScreenUpdating = False
- 关闭自动计算：Application.Calculation = xlCalculationManual
- 用数组批量读写：arr = Range.Value → 处理 → Range.Value = arr
- 操作完成后恢复设置

**验证标准**：
- 抽样检查 3-5 个位置确认修改正确
- 检查错误日志中是否有失败记录
- 确认 ScreenUpdating 和 Calculation 已恢复`;
}

/**
 * 条件格式与数据验证场景
 *
 * 核心原则：规则可读、不冲突、可追溯
 */
function scenarioConditionalFormat(): string {
  return `### 条件格式与数据验证

**触发场景**：用户要求高亮单元格、添加下拉列表、设置输入规则、数据条/色阶等

**标准流程**：
1. 用 selection.get 确认目标区域
2. 用 script.execute 调用条件格式/数据验证 API
3. 用 range.read 确认规则已生效

**质量规则**：
- 条件格式公式必须引用区域左上角单元格为锚点（如 B2 区域的公式写 \`=B2>100\`）
- 多条规则时注意优先级和"如果为真则停止"设置，避免规则冲突
- 数据验证的提示信息必须写明输入要求（"请输入 0-100 之间的整数"）
- 色阶/数据条优先于图标集（更直观、不依赖图例）

**禁止项**：
- ❌ 整列条件格式（如 A:A）— 性能极差，应限定实际数据范围
- ❌ 条件格式公式引用偏移（区域从 B2 开始但公式写 \`=A2>100\`）
- ❌ 超过5条条件格式规则在同一区域 — 合并规则或简化逻辑
- ❌ 数据验证不设错误提示（用户不知道为什么输入被拒绝）

**常用规则**：
- 高亮大于/小于阈值 → FormatConditions(xlCellValue)
- 根据公式条件着色 → FormatConditions(xlExpression)
- 数据条 → FormatConditions.AddDatabar
- 色阶 → FormatConditions.AddColorScale
- 下拉列表 → Validation.Add Type:=xlValidateList, Formula1:="选项1,选项2"

**验证标准**：
- 确认 FormatConditions.Count 或 Validation.Type 符合预期
- 抽样测试：输入应触发/不应触发规则的值，确认行为正确`;
}

/**
 * 数据建模与预测场景
 *
 * 核心原则：假设明确、方法透明、结果带置信度
 */
function scenarioModeling(): string {
  return `### 数据建模与预测

**触发场景**：用户要求回归分析、趋势预测、分类、聚类等

**标准流程**：
1. 用 range.read 读取数据，评估数据量和变量类型
2. 用 script.execute 实现统计算法（线性回归/移动平均/指数平滑等）
3. 将模型参数和预测结果写入工作表
4. 输出模型摘要和预测区间

**质量规则**：
- 必须声明模型假设（线性关系、正态分布、独立性等）
- 预测结果必须带置信区间或误差范围，不能只给点估计
- 时间序列预测必须标注预测起点（"2024-07 之后为预测值"）
- 模型拟合指标必须报告（R²、RMSE、MAE 等）

**禁止项**：
- ❌ 外推预测不标注不确定性（"明年销售额将是 500 万" → 应为"预测 500 万 ± 80 万"）
- ❌ 忽略数据前提直接建模（时序数据有趋势/季节性却用简单均值）
- ❌ 过拟合不报告（10个数据点用5个参数的模型）
- ❌ 将预测结果混入原始数据不标注

**验证标准**：
- 模型参数合理性检查（回归系数符号是否符合业务逻辑）
- 预测值量级合理性（不能出现负数销售额等）
- 回读确认预测结果已正确写入`;
}

/**
 * 文件操作场景
 *
 * 核心原则：路径必须绝对、先查路径再操作、多簿切换清晰
 */
function scenarioFileOps(): string {
  return `### 文件操作

**触发场景**：用户要求创建新文件、打开已有文件、保存/另存为、在多个工作簿间切换

**标准流程**：
1. 如用户提到"桌面""文档""下载"等位置，先用 file.getPaths 获取绝对路径
2. 涉及 Excel 创建、编辑、保存、验证前，先调用 office.connection.status({ app:"excel" }) 检测连接状态
3. 创建/编辑磁盘上的 Excel 文件 → 优先用 office.action.apply（Open XML 内置能力）→ office.action.validate 确认；只有返回 unsupported/needsCom/failed 后才使用 python.execute 或其它兜底
4. 创建非 Excel 文件 → shell.execute({ command: "..." }) 用系统命令创建
5. 用户明确要操作当前打开的 Excel 窗口/选区，且 office.connection.status 显示已连接 → workbook.open/workbook.inspect/selection.get/range.* → workbook.save
6. 多工作簿 → 用 workbook.switch(workbookName) 切换 → 操作 → 切回
7. 另存为 → workbook.save({ saveAsPath: "新路径" })

**工具选择规则**：
- 文件级创建/编辑 .xlsx → 优先 office.action.apply，不要先用 openpyxl、pip install 或 shell 拼 Python
- 已连接 Excel + 操作当前窗口/选区 → 优先 selection.get、range.*、workbook.* 等专用工具；复杂批处理可用 python.execute，失败后再兜底
- 未连接 Excel + 创建/编辑 .xlsx → 必须先用 office.action.* 的 Open XML 能力；失败后再用 python.execute 或其它兜底
- 创建非 Excel 文件（.txt/.csv/.json 等）→ 用 shell.execute
- 运行系统命令（Git/pip/dir 等）→ 用 shell.execute
- Excel 内自动化脚本 → 用 script.execute（自动注入 COM 连接）

**质量规则**：
- filePath 必须是绝对路径，不能用相对路径或仅文件名
- Windows 路径用反斜杠：C:\\Users\\用户\\Desktop\\报表.xlsx
- 创建工作簿时，如果目录不存在应先告知用户
- 保存前确认修改已完成，避免保存半成品
- 多工作簿操作时，每次切换前确认目标工作簿名称（含扩展名）
- shell.execute 的 workdir 默认为用户主目录，操作特定目录时必须指定

**禁止项**：
- ❌ 猜测用户路径（如假设桌面在 C:\\Users\\Admin\\Desktop）→ 必须用 file.getPaths
- ❌ 不保存就结束（修改后必须 workbook.save）
- ❌ 覆盖已有文件不提醒（workbook.create 到已存在的路径会覆盖）
- ❌ 用 shell.execute 执行 Excel 自动化脚本（应使用 script.execute，自动注入 COM 连接）
- ❌ 用 shell.execute 拼 python -c 或 echo ... | python 执行多行脚本（应使用 python.execute）
- ❌ shell.execute 中执行危险系统命令（格式化磁盘、删除系统文件等）

**验证标准**：
- 创建/打开后用 workbook.inspect 确认工作簿已激活
- 保存后文件路径可访问
- 多簿切换后用 workbook.inspect 确认当前活动工作簿正确
- shell.execute 后检查 exitCode 是否为 0，stderr 是否有错误`;
}

export function buildScenarioPromptSection(): string {
  return [
    scenarioCodeQA(),
    scenarioDataCleaning(),
    scenarioChart(),
    scenarioReport(),
    scenarioBatch(),
    scenarioConditionalFormat(),
    scenarioModeling(),
    scenarioFileOps(),
  ].join("\n");
}
