# 自动化基础层

职责：承载可复用的本机自动化基础能力，例如 PowerShell、COM、Python、进程执行、变量编码和 JSON 解析。

模块说明：

- `powershell.ts`: PowerShell 执行、UTF-8 输出设置和字符串变量安全注入。
- `python.ts`: Python 运行时定位、xlwings 上下文注入和脚本执行。
- `scriptEngine.ts`: 优先使用 Python/xlwings，失败时回退 PowerShell COM。
- `json.ts`: 解析自动化脚本返回的 JSON 输出。

关联模块：

- `../tools/implementations`: Office/Excel bridge 复用本层能力。
- `../security/sandbox`: Shell 工具的通用进程安全策略不放在本层。
