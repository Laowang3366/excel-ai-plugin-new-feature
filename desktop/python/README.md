# 嵌入式 Python 环境

此目录存放嵌入式 Python 发行版（Python Embedded Distribution），
用于在不安装 Python 的用户机器上提供 Python + xlwings 支持。

## 目录结构（安装后）

```
python/
├── python.exe          # Python 解释器
├── python3.dll         # 运行时 DLL
├── python311.dll       # 版本特定 DLL
├── python311.zip       # 标准库（精简版）
├── Lib/                # 第三方包
│   └── site-packages/
│       └── xlwings/    # Excel 互操作库
└──Scripts/
    └── xlwings.exe
```

## 安装步骤

运行 `setup-python-embed.ps1` 脚本自动下载和配置：

```powershell
cd desktop
powershell -ExecutionPolicy Bypass -File python\setup-python-embed.ps1
```

脚本会：
1. 下载 Python 3.11 Embedded Distribution (~10MB)
2. 配置 pip（解压标准库 + 安装 get-pip.py）
3. 安装 xlwings（`pip install xlwings`）
4. 验证 `python.exe -c "import xlwings"` 成功

## 注意事项

- 此目录不提交到 Git（已在 .gitignore 中排除）
- 构建安装包时，electron-builder 会将此目录打包为 extraResources
- 如果用户机器已有系统 Python + xlwings，会优先使用系统 Python
- 如果此目录不存在，应用会 fallback 到 JScript(cscript.exe) 或 PowerShell

## 体积优化

精简后预计约 30-40MB，可通过以下方式进一步压缩：
- 删除 `Lib/site-packages` 中不需要的包
- 使用 UPX 压缩 DLL
- 只保留 xlwings 核心模块，移除测试/文档
