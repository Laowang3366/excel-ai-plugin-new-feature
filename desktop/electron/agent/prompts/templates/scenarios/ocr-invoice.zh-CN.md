## 场景化操作指南：OCR 与发票识别

当本轮涉及图片/PDF/Office 可见内容解析、OCR、识别字段、发票识别，或模型没有多模态能力但需要理解图片内容时：
- 第一步调用 `ocr.parseDocument`，图片/PDF/Office 文件路径来自本轮附件或用户给出的路径。
- 发票场景必须使用 `ocr.parseDocument({ mode:"invoice", filePaths:[...] })`，不要只凭文件名或历史对话作答。
- 基于 OCR 返回的 text、markdown、rows、warnings、fallbacks 抽取字段；缺失字段填空字符串，禁止编造。
- 发票默认字段：文件名、发票类型、发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、校验码、备注。
- 多张发票按一张一行整理，第一行为字段名；需要写入 Excel/WPS 时先检查连接和选区，再用 `range.write` 写入，写入后回读验证一次。
- 普通图片、PPT 界面、Word 文档、Excel 样式美化等视觉判断任务，也可以先用 `ocr.parseDocument` 获取可见内容，再选择 Office 工具修改或给出判断。
