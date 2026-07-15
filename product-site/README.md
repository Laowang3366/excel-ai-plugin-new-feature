# 文格 AI 助手产品站

独立产品页、安装包分发、更新清单和下载统计后台。服务默认只监听 `127.0.0.1:18120`，由 Nginx 对外提供 `plugin.shelelove.top`。

## 本地运行

```bash
npm install
npm start
```

- 产品页：`http://127.0.0.1:18120/`
- 后台：`http://127.0.0.1:18120/admin/`
- 健康检查：`http://127.0.0.1:18120/healthz`
- 本地默认后台密码仅用于开发：`development-admin`

生产环境必须配置 `.env.example` 中的密码哈希、Cookie 密钥、统计盐值和数据治理周期。默认只保留 90 天下载记录，IP 匿名标识每 30 天按 UTC 周期轮换；UA 最多保留 200 字符，Referer 只保留站点 Origin。密码哈希通过以下命令生成：

```bash
npm run hash-password -- "your-long-password"
```

## 发布安装包

`publish-release` 会复制 NSIS 安装包、blockmap 和 `latest.yml`，计算 SHA-256，并使用 Ed25519 私钥签署桌面端更新清单：

```bash
npm run publish-release -- \
  --version 0.1.81 \
  --installer ../desktop/release/Wengge-AI-Assistant-Setup-0.1.81.exe \
  --blockmap ../desktop/release/Wengge-AI-Assistant-Setup-0.1.81.exe.blockmap \
  --latest-yml ../desktop/release/latest.yml \
  --notes-file ../release-notes/0.1.81.json \
  --private-key ../desktop/.secrets/update-private.pem \
  --output ./.local/releases \
  --base-url https://plugin.shelelove.top
```

私钥只允许存放在发布机或服务器的受限目录，不得提交 Git。桌面安装包只包含对应公钥。

## 热补丁范围

桌面端只接受以下路径，其他文件会在解压前被拒绝：

- `dist/**`
- `public/knowledge/**`
- `public/wps-jsa-bridge/**`

主进程、preload、原生依赖和 Python 运行时必须发布完整安装包。创建补丁：

```bash
cd ../desktop
npm run patch:build -- --id patch-001 --base-version 0.1.81 --output release/patch-001.zip
```

## 生产目录

```text
/opt/wenge-product/app       产品站代码
/opt/wenge-product/data      SQLite 下载统计
/opt/wenge-product/releases  安装包、latest.yml 与签名清单
/opt/wenge-product/secrets   更新签名私钥
/etc/wenge-product.env       运行环境变量
```

部署使用独立系统用户、独立 systemd 服务、独立 Nginx 配置和日志，不依赖或修改服务器上的旧 Excel 项目容器。
