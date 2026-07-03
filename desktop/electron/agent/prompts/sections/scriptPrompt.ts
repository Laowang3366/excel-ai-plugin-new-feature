/**
 * 脚本规范提示词：Python、JavaScript、VBA 的注入变量和代码模板。
 *
 * 关联模块：
 * - ../systemPrompt.ts: 组装完整系统提示词。
 * - ../systemPrompt.test.ts: 校验关键提示词内容不丢失。
 */

// ============================================================
// 脚本代码规范
// ============================================================

export function scriptSpec(): string {
  return `## 脚本代码规范

### Python（首选，所有环境通用）
自动注入变量：app(xw.apps.active)、wb(app.books.active)、ws(wb.sheets.active)
\`\`\`python
# 设置单元格值
ws.range("A1").value = "Hello"
# 读取并计算
val = ws.range("B1").value
ws.range("C1").value = val * 2
# 批量操作
ws.range("A1:A10").value = [[i] for i in range(1, 11)]
# 输出结果
print(json.dumps({"result": "ok"}, ensure_ascii=False))
\`\`\`

### JavaScript（cscript.exe，Windows 内置）
自动注入变量：excel(Application)、wb(ActiveWorkbook)、ws(ActiveSheet)
\`\`\`javascript
// 设置单元格值
ws.Range("A1").Value = "Hello";
// 读取并计算
var val = ws.Range("B1").Value;
ws.Range("C1").Value = val * 2;
// 输出结果
WScript.Echo(JSON.stringify({result: "ok"}));
\`\`\`

### VBA（最终兜底）
必须包含 Sub Main() ... End Sub 入口
\`\`\`vba
Sub Main()
    Range("A1").Value = "Hello"
End Sub
\`\`\``;
}
