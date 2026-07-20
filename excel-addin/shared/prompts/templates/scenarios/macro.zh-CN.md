## 宏 / VBA / JSA 边界（本加载项）

- 本加载项**不实现** `macro.detect` / `macro.write` / `macro.run`，也不能向工作簿写入或运行 VBA / WPS JSA 内部宏。
- 不要只输出宏代码并声称“已安装/已绑定按钮/已运行”；不得调用或编造 `ui.addControl` / `ui.listControls` / ActiveX / UserForm。
- 可用替代：用 `range.write` / `formula.write` 与已注册工作表/表格/图表工具完成数据与展示层目标；若用户坚持宏方案，明确说明需在桌面端完整应用或宿主宏编辑器中手工完成，且本加载项无法代为执行。
- 公式、数值、文本与格式仍走 `range.*` / `formula.*`，不得用“宏已执行”冒充单元格变更。
