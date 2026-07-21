## 场景化操作指南：OCR 与发票识别（Excel 加载项）

当本轮涉及图片 OCR、字段抽取、发票识别，或用户消息包含图片多模态内容时：

- 直接阅读本轮用户消息中的图片内容；禁止工具 `ocr.parseDocument`（加载项无此工具）；不要索取磁盘路径；不要假设 MinerU/Electron IPC。
- 发票场景基于可见票据内容抽取字段；缺失字段填空字符串，禁止编造。
- 发票默认字段：文件名、发票类型、发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、校验码、备注。
- 多张发票按一张一行整理；`invoices` 数组一项一票，`fields` 使用上述字段名。
- **PDF**：若本轮仅有 PDF 或宿主/模型无法可靠解析 PDF，返回 typed unsupported 说明，禁止假成功。
- 识别结果中的 Base64 / API Key 不得回显给用户。
- **不要**在本轮调用 `range.write`：加载项 UI 会解析结构化结果并在用户确认后经审批写入。
- **结构化输出（强制）**：回复末尾必须包含且仅包含一处：

<<<WENGGE_OCR_RESULT_V1
{"kind":"invoice","text":"全文摘要","fields":{},"rows":[],"invoices":[{"filename":"a.png","text":"","fields":{"发票号码":"…"},"rows":[]}],"errors":[]}
WENGGE_OCR_RESULT_V1>>>

`kind` 为 `image` 或 `invoice`。通用图片可把正文放在 `text`，表格放 `rows`（首行表头）或 `fields`。标记外可写简短说明；解析失败时 UI 只展示原始文本。
