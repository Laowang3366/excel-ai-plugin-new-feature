# WPS JSA 加载说明

本目录是 **独立验证用** WPS 加载项清单，不是 `desktop/public/wps-jsa-bridge` 的 Electron localhost 桥。

## 与桌面桥的区别

| | 本加载项 | desktop WPS JSA bridge |
|---|---|---|
| 目标 | 任务窗格内直接调用 `Application` | Worker 经 `127.0.0.1:45221` 轮询写宏 |
| 能力 | selection/range/formula/sheet（有运行时探测） | detect / CodeModule write only |
| 依赖 | 无 COM/.NET/Electron | .NET Worker + token |

## 安装（开发）

1. 在 `excel-addin/` 执行 `npm run dev`，确认 `https://localhost:3000` 可访问（默认 HTTPS；仅调试用 `npm run dev:http` 时才是 HTTP）。
2. 将本目录 `manifest.xml` 与指向任务窗格 URL 的入口页按 WPS JS 加载项规范放入用户 `jsaddons` 目录，或使用 WPS 开发者工具侧载。
3. **完整重启 WPS** 后再打开表格。
4. 打开任务窗格后查看「宿主状态」：应显示 `wps-jsa`。

## 合同假设

仓库内可核实的 JSA 合同仅有 `Application` / `ActiveWorkbook` / 宏 CodeModule。
区域与工作表 API 按常见 ET 对象模型实现；若成员缺失，工具返回 typed `unsupported`，不伪造成功。

## 不支持

- 桌面 COM Worker 协议
- 文件级 Open XML
- Power Query / 透视表 / 图表等后续批次能力（见 `docs/capability-matrix.md`）

## 正式本地 jsaddons 包

仓库可生成**正式本地 file:// 包**（非下一阶段占位）。**真实 WPS 侧载尚未验收**。

```bash
# 在 excel-addin/ 下
npm run manifest:wps:check
npm run package:wps -- --git-sha <sha>
```

产物 `dist/` 布局（对齐桌面 `WpsJsaService` / `public/wps-jsa-bridge` 合同）：

| 路径 | 说明 |
|------|------|
| `publish.xml` | jsaddons 根级注册；`type=et`，`url=file://%AppData%/kingsoft/wps/jsaddons/WenggeExcelAiAddin_/index.html` |
| `WenggeExcelAiAddin_/` | 任务窗格静态资源（`--base ./`）、`manifest.xml`、`ribbon.xml`、`wps-entry.js` |
| `BUILD_INFO.json` / `SHA256SUMS.txt` | 构建元数据与哈希 |

Windows 安装骨架（需完整重启 WPS；本仓库未验收）：

1. 将 `WenggeExcelAiAddin_/` 复制到 `%AppData%\kingsoft\wps\jsaddons\`。
2. 将 `publish.xml` 中对应 `jsplugin` 条目合并进同目录下的 `publish.xml`（若已有其它加载项，保留其它节点）。
3. 完整退出并重启 WPS 表格后再打开任务窗格，宿主状态应为 `wps-jsa`。

本包为构建期脚本产出；加载项运行时仍不引入 COM/.NET/Electron/`child_process`。
