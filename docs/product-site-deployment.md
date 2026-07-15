# 产品站部署

## 生产拓扑

```text
plugin.shelelove.top:443
  -> Nginx
     -> /releases/windows/*  静态安装包与 electron-updater 元数据
     -> /download/windows    Fastify 记录统计后 X-Accel-Redirect
     -> 其他请求             127.0.0.1:18120
```

服务目录：

```text
/opt/wenge-product/app
/opt/wenge-product/data
/opt/wenge-product/backups
/opt/wenge-product/releases
/opt/wenge-product/secrets
```

服务使用独立系统用户 `wenge-product`，systemd 单元为 `wenge-product.service`。Nginx 配置为 `/etc/nginx/conf.d/plugin.shelelove.top.conf`，日志为 `/var/log/nginx/wenge-product.*.log`。

## 隔离原则

- 不停止、删除或修改服务器已有 Docker 容器。
- 不占用已有的 8787、18080-18082、18789、18888、19899 等端口。
- 不修改其他域名的 Nginx 文件和证书。
- 上线前必须执行 `nginx -t`；失败时不 reload。
- 应用只监听 `127.0.0.1:18120`，公网只开放 Nginx 80/443。

## 运行配置

生产环境变量存放在 `/etc/wenge-product.env`，权限 `0600`。必须设置：

- `ADMIN_PASSWORD_HASH`
- `COOKIE_SECRET`
- `ANALYTICS_SALT`
- `ANALYTICS_RETENTION_DAYS=90`（允许 1-3650 天）
- `ANALYTICS_IP_ROTATION_DAYS=30`（允许 1-365 天）
- `ANALYTICS_BACKUP_DIR=/opt/wenge-product/backups`
- `ANALYTICS_BACKUP_RETAIN=14`（允许 1-365 份）
- `DATA_DIR`
- `RELEASES_DIR`
- `DATABASE_PATH`
- `USE_ACCEL_REDIRECT=true`

## 下载统计数据治理

- 不保存原始 IP；服务使用 `ANALYTICS_SALT` 派生周期密钥，按 UTC 周期生成不可逆 HMAC 标识。
- 同一 IP 在一个轮换周期内保持一致，跨周期会生成不同标识。因此跨周期查询的“独立下载”可能重复计数，这是降低长期关联能力的预期取舍。
- 下载记录默认保留 90 天。服务启动时立即清理过期记录，并每 6 小时再次清理；清理定时器不阻塞进程退出。
- 清理失败只写告警，不阻塞新的统计写入或安装包下载。生产监控应对 `download analytics maintenance failed` 告警。
- UA 最多保存 200 字符；Referer 只保存 `scheme://host[:port]`，不保留路径、查询参数或片段。首次升级会对既有记录执行同样的数据最小化迁移。

## 备份、恢复与演练

当前运维基线目标为 RPO 不超过 25 小时、RTO 不超过 4 小时：每天 03:15 执行一次在线一致性备份，并增加最多 30 分钟随机延迟；默认保留最近 14 份。此目标需要由生产负责人确认，且不替代异机/异地副本。

安装定时单元：

```bash
install -d -o wenge-product -g wenge-product -m 0700 /opt/wenge-product/backups
install -o root -g root -m 0644 deploy/wenge-product-backup.service /etc/systemd/system/
install -o root -g root -m 0644 deploy/wenge-product-backup.timer /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now wenge-product-backup.timer
systemctl start wenge-product-backup.service
systemctl status wenge-product-backup.service --no-pager
```

备份通过 SQLite 在线备份 API 生成，不直接复制活动 WAL 文件。归档切换为自包含的 `DELETE` journal 模式，并带有包含大小和 SHA-256 的同名 `.json` 元数据。生成、独立验证和恢复前均执行 `quick_check`，并确认 `downloads` 表存在。

季度恢复演练或真实恢复时，先选择备份并验证，再停止应用、隔离当前数据库及其 WAL/SHM，最后恢复到空目标路径：

```bash
cd /opt/wenge-product/app
npm run analytics:verify -- --backup /opt/wenge-product/backups/analytics-<timestamp>.sqlite
systemctl stop wenge-product.service
install -d -o wenge-product -g wenge-product -m 0700 /opt/wenge-product/data/quarantine-<timestamp>
for file in /opt/wenge-product/data/analytics.sqlite /opt/wenge-product/data/analytics.sqlite-wal /opt/wenge-product/data/analytics.sqlite-shm; do
  [ ! -e "$file" ] || mv -- "$file" /opt/wenge-product/data/quarantine-<timestamp>/
done
npm run analytics:restore -- --backup /opt/wenge-product/backups/analytics-<timestamp>.sqlite --target /opt/wenge-product/data/analytics.sqlite
chown wenge-product:wenge-product /opt/wenge-product/data/analytics.sqlite
systemctl start wenge-product.service
curl -fsS http://127.0.0.1:18120/healthz
```

每次演练需记录备份时间、恢复开始/结束时间、`quick_check` 结果、后台统计抽样结果和实际 RPO/RTO。`wenge-product-backup.service` 失败必须接入 systemd/日志告警；当前仓库只提供可监控的失败退出码，告警接收端仍需在生产平台配置。

## 验收

```bash
curl -fsS http://127.0.0.1:18120/healthz
curl -fsS https://plugin.shelelove.top/healthz
curl -I https://plugin.shelelove.top/download/windows
systemctl status wenge-product --no-pager
nginx -t
```

还需要在浏览器验证产品页、后台登录、统计周期切换，以及实际下载后后台计数增加。
