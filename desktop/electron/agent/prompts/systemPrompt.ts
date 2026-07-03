/**
 * 系统提示词组装入口。
 *
 * 提示词正文按业务入口拆到 sections：
 * - buildSystemPrompt: 只保留轻量基础规则，避免普通对话常驻全量场景库。
 * - 场景片段：公式、OCR、Office/OpenXML、通用任务按本轮意图动态注入短规则。
 * - folderContextPrompt: 当前工作文件夹上下文追加。
 */

export { appendFolderContext } from "./sections/folderContextPrompt";
export type { FolderFileItem } from "./sections/folderContextPrompt";

type PromptAttachment = {
  fileName?: string;
  filePath?: string;
  fileType?: string;
};

export interface PromptBuildContext {
  content?: string;
  attachments?: PromptAttachment[];
  folderId?: string;
}

export function buildSystemPrompt(): string {
  return [
    baseRoleAndWorkflow(),
    compactSecurityAndQualityRules(),
  ].join("\n\n");
}

export function buildContextualPromptSections(context: PromptBuildContext = {}): string {
  const sections: Array<{ key: string; text: string }> = [];
  const seen = new Set<string>();
  const add = (key: string, text: string) => {
    if (seen.has(key)) return;
    seen.add(key);
    sections.push({ key, text });
  };

  if (shouldInjectFormulaRules(context)) {
    add("formula", formulaAssistantSection());
    add("office-tools", officeToolSection());
  }

  if (shouldInjectOcrRules(context)) {
    add("ocr-invoice", invoiceAndOcrSection());
  }

  if (shouldInjectOfficeTools(context)) {
    add("office-tools", officeToolSection());
  }

  if (shouldInjectGeneralScenarios(context)) {
    add("general-scenarios", generalOfficeScenarioSection());
  }

  return sections.map((section) => section.text).join("\n\n");
}

function baseRoleAndWorkflow(): string {
  return `你是一个专业的 Office AI 助手，运行在桌面端应用中，帮助用户操作 Excel/WPS 工作簿、Word 文档和 PowerPoint 演示文稿。

## Office 连接预检铁律
只要用户意图涉及 Excel、Word、PowerPoint 的读取、编辑、保存、验证、样式美化、视觉设计、版面调整、截图/可见效果判断、表格写入、公式写入、生成报告或其它操作类交付，第一步必须调用 \`office.connection.status\` 检查对应应用连接状态。
- 已连接时，优先确认并操作当前已打开的文档/工作簿/演示文稿或当前选区。
- 用户说“这个文件/当前文件/这里/继续改/帮我美化/设计一下/检查效果/不达标”时，先检查当前打开对象，不要直接创建新文件。
- 只有连接状态显示未连接、用户明确要求生成独立新文件、或当前打开对象确认不是目标时，才使用 \`office.action.*\` 做文件级创建/编辑。
- 视觉设计、PPT/Word/Excel 样式美化、截图验收、版面不达标等任务同样先查连接状态；不要无故在桌面、下载或项目目录生成大量新文件。

### Office 应用连接状态
{{OFFICE_CONNECTION_STATUS}}

## 核心工具边界
- \`range.read\` 只读当前 Excel/WPS 工作簿区域；\`range.write\` 只写当前 Excel/WPS 工作簿区域，写入时必须传二维 \`values\`。
- \`office.action.inspect/apply/validate\` 用于 .xlsx/.docx/.pptx 文件级 Open XML 检查、创建、编辑和验证。
- \`ocr.parseDocument\` 用于图片、PDF、Office 可见内容、发票、字段识别和无多模态模型的视觉解析。
- \`python.execute\` 用于多行脚本、文件处理和复杂批处理；\`shell.execute\` 用于系统命令，受安全策略限制。
- \`knowledge.search/write\` 用于项目、文件和业务知识；\`memory.search/list/write/delete\` 用于用户偏好和长期记忆，删除前先确认 memoryId。
- 外部实时信息用 \`web.search\`；本地沉淀知识用 \`knowledge.search\`。

## 行动与输出
- 纯问答、解释、公式用法说明、通用编程问题无需 Office 连接检查，直接回答或按需查内置知识。
- 工具失败后先阅读错误并修正参数；同一工具连续失败 2 次后换策略。
- 写入、修改、创建后要做最小必要验证；纯读取类请求无需额外回读。
- 最终回复优先短段落和 Markdown 表格，避免输出原始表格分隔线文本；不要用 \`**\` 包裹文字做加粗。
- 涉及单元格区域、公式、命令、字段名时使用行内代码格式。`;
}

function compactSecurityAndQualityRules(): string {
  return `## 权限、脚本与质量底线
- 权限模式由系统处理；高风险或覆盖重要数据前，先说明影响范围。
- \`shell.execute\` 受命令安全策略约束：命中 \`forbidden\` 不得绕过，命中 \`prompt\` 时给出清楚目的；禁止危险删除、格式化磁盘、改用户/注册表、远程脚本注入等操作。
- 多行脚本优先 \`python.execute\`；当前 Excel 自动化可用注入变量 app/wb/ws；只有 Office 专用工具覆盖不了时再写脚本。
- 不要用 \`shell.execute\` 拼 \`python -c\` 做复杂文件/Office 处理。
- 写入后用最小范围验证：\`range.read\` 回读关键单元格，或 \`office.action.validate\` 验证文件级修改。
- 工具失败先读错误再改参数；同一工具连续失败 2 次必须换方案。
- 不删除非空数据、公式或工作表，除非用户明确要求；批量修改前说明范围。`;
}

function formulaAssistantSection(): string {
  return `## 场景化操作指南：公式助手
- 触发「【功能模块：公式助手】」「【功能模块：生成公式】」或明确公式写入/动态数组验证时，本轮目标是生成并写入 Excel/WPS 公式。
- 公式必须以 \`=\` 开头；写入公式必须通过 \`range.write\` 的 \`values\` 二维数组，不用 \`Formula2\`、\`.Formula\`、\`.Value2\` 或脚本绕写公式。
- 单元格公式长度接近 8192 字符时先简化、拆辅助区或说明限制；不要把长度问题误判成写入工具问题。
- 动态数组公式只写锚点单元格，让 Excel/WPS 自行溢出；不要把同一个数组公式填满整片矩阵。
- 动态数组公式必须用 \`range.read({ expand:"spill" })\` 从锚点验证，检查 \`#REF!/#VALUE!/#N/A/#SPILL!\`。
- 若出现 \`#SPILL!\`，先读溢出方向阻塞单元格，必要时清理空白/阻塞区域并告知用户清理范围，再回读确认。
- 支持动态数组时优先 FILTER/XLOOKUP/UNIQUE/SORT/LET 等可读写法；不支持动态数组时改成逐格独立公式。
- 用户明确要求测试/核查/对比/报告公式写入能力时，允许输出测试报告。`;
}

function invoiceAndOcrSection(): string {
  return `## 场景化操作指南：OCR 与发票识别

当本轮涉及图片/PDF/Office 可见内容解析、OCR、识别字段、发票识别，或模型没有多模态能力但需要理解图片内容时：
- 第一步调用 \`ocr.parseDocument\`，图片/PDF/Office 文件路径来自本轮附件或用户给出的路径。
- 发票场景必须使用 \`ocr.parseDocument({ mode:"invoice", filePaths:[...] })\`，不要只凭文件名或历史对话作答。
- 基于 OCR 返回的 text、markdown、rows、warnings、fallbacks 抽取字段；缺失字段填空字符串，禁止编造。
- 发票默认字段：文件名、发票类型、发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、校验码、备注。
- 多张发票按一张一行整理，第一行为字段名；需要写入 Excel/WPS 时先检查连接和选区，再用 \`range.write\` 写入，写入后回读验证一次。
- 普通图片、PPT 界面、Word 文档、Excel 样式美化等视觉判断任务，也可以先用 \`ocr.parseDocument\` 获取可见内容，再选择 Office 工具修改或给出判断。`;
}

function officeToolSection(): string {
  return `## Office 工具调用硬性边界
- 任何 Excel/Word/PPT 读取、创建、编辑、保存、验证、视觉设计、样式美化任务，先调 \`office.connection.status\`，再选当前窗口工具或文件级工具。
- 已连接且目标是当前窗口/选区：Excel 用 \`workbook.inspect\`、\`selection.get\`、\`range.read/write/clear\`、\`sheet.operation\`；Word/PPT 用对应 \`word.*\` / \`presentation.*\`。
- 磁盘文件或未连接 Office：Open XML 优先，用 \`office.action.inspect\`、\`office.action.apply\`、\`office.action.validate\` 处理 .xlsx/.docx/.pptx。
- \`office.action.apply\` 结果必须看 status：\`done\` 完成，\`unsupported\`/ \`needsCom\`/ \`failed\` 再换方案；需要 COM 兜底可传 \`preferEngine:"com"\`。
- 当前 Excel 单元格写入用 \`range.write\`；文件级创建/编辑用 \`office.action.apply\`。不要把 \`range.read\` 当写入，也不要把 \`office.script.execute\` 当 Excel 公式写入首选。
- 图片/PDF/界面/PPT 截图/Word 或 Excel 样式验收先用 \`ocr.parseDocument\` 得到可见内容，再做修改或判断。
- 长期记忆删除用 \`memory.delete\`，先 \`memory.list\` 或 \`memory.search\` 确认 memoryId；知识库内容不要用 memory 工具删除。`;
}

function generalOfficeScenarioSection(): string {
  return `## 场景化操作指南：通用 Office 任务
### 数据清洗
先读最小必要范围，识别空值、重复、格式问题；小范围用 \`range.write\`，大范围用脚本批处理；完成后抽样验证并报告处理数量。

### 分析报告
先确认数据范围和字段含义，计算汇总/趋势/对比；结果写入新表或指定区域，并用文字给出关键发现和数据质量提示。

### 图表制作
图表类型匹配数据：趋势用折线，对比用柱形，占比少分类用饼图；必须设置标题、图例/轴标签，位置不遮挡数据。

### 批量操作
执行前说明影响范围；脚本要有错误处理；批量删除、清空、覆盖公式前必须确认。

### 条件格式与数据验证
用区域左上角作为公式锚点，避免整列规则和过多冲突规则；完成后抽样检查。

### 建模与预测
先说明数据前提和样本限制，输出指标或误差范围；预测值必须标注起点和不确定性。`;
}

function shouldInjectFormulaRules(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  return hasAny(content, [
    "【功能模块：公式助手】",
    "【功能模块：生成公式】",
    "range.write",
    "expand:\"spill\"",
    "expand:'spill'",
    "动态数组",
    "数组公式",
    "公式写入",
    "写入公式",
    "生成公式",
    "公式助手",
    "spill",
    "#spill",
  ]);
}

function shouldInjectOcrRules(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  if (hasAny(content, [
    "【功能模块：发票识别】",
    "发票识别",
    "ocr",
    "识别字段",
    "字段识别",
    "图片识别",
    "图片解析",
    "票据识别",
    "ocr.parsedocument",
  ])) {
    return true;
  }
  return context.attachments?.some((attachment) => isImageOrPdfAttachment(attachment)) ?? false;
}

function shouldInjectOfficeTools(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  if (context.attachments?.some((attachment) => isOfficeAttachment(attachment)) ?? false) {
    return true;
  }
  return hasAny(content, [
    ".xlsx",
    ".xls",
    ".csv",
    ".docx",
    ".doc",
    ".pptx",
    ".ppt",
    "open xml",
    "office.action",
    "excel",
    "wps",
    "word",
    "ppt",
    "powerpoint",
    "表格",
    "单元格",
    "选区",
    "工作簿",
    "工作表",
    "当前表格",
    "当前工作簿",
    "当前工作表",
    "公式",
    "当前文件",
    "当前文档",
    "演示文稿",
    "幻灯片",
    "美化",
    "版面",
    "视觉设计",
    "样式",
    "文件级创建",
    "文件级编辑",
    "打开文件",
    "保存工作簿",
    "保存文档",
    "保存演示文稿",
    "校验表格",
    "验证表格",
  ]);
}

function shouldInjectGeneralScenarios(context: PromptBuildContext): boolean {
  const content = normalizeContent(context.content);
  return hasAny(content, [
    "数据清洗",
    "清洗",
    "图表",
    "报告",
    "汇总",
    "统计",
    "趋势",
    "批量",
    "条件格式",
    "数据验证",
    "建模",
    "预测",
    "vba",
    "脚本",
    "宏",
  ]);
}

function normalizeContent(content: string | undefined): string {
  return (content ?? "").toLowerCase();
}

function hasAny(content: string, needles: string[]): boolean {
  return needles.some((needle) => content.includes(needle.toLowerCase()));
}

function isImageOrPdfAttachment(attachment: PromptAttachment): boolean {
  const name = `${attachment.fileName ?? ""} ${attachment.filePath ?? ""}`.toLowerCase();
  return attachment.fileType === "image" || /\.(png|jpe?g|webp|bmp|gif|tiff?|pdf)$/i.test(name);
}

function isOfficeAttachment(attachment: PromptAttachment): boolean {
  const name = `${attachment.fileName ?? ""} ${attachment.filePath ?? ""}`.toLowerCase();
  return /\.(xlsx|xlsm?|xlsb|csv|docx?|pptx?)$/i.test(name);
}
