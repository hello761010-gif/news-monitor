# News Monitor V6 — GitHub Actions 部署包

本包用于在 **Public GitHub Repository** 中定时收集香港与台湾四家媒体的 RSS 新闻。工作流每小时运行一次，也可以在 Actions 页面手动运行。

## 文件结构

```text
news-monitor/
├── news_monitor_v6.py
├── requirements_news_monitor_v6.txt
├── config.example.json
├── .gitignore
└── .github/
    └── workflows/
        └── news-monitor.yml
```

## 1. 上传文件

将本包中的文件放到仓库根目录，必须保留 `.github/workflows/news-monitor.yml` 这个路径。`news_monitor_v6.py` 和 `requirements_news_monitor_v6.txt` 必须位于仓库根目录。

不要上传真实 `config.json`、`news_monitor.db`、日志或包含新闻正文的 CSV。仓库中的 `.gitignore` 可避免本地使用 Git 时误提交这些文件，但不会自动删除已经提交过的文件。

## 2. 设置 GitHub Repository Secrets

进入：`Settings → Secrets and variables → Actions → New repository secret`。

### Secret A: `NEWS_MONITOR_CONFIG`

将你原来项目中真实的 `config.json` 全文作为 Secret 值。必须包含 `sources` 对象，且四家媒体各自的 `feeds` 数组要有真实可访问的 RSS 地址。

请在原配置中确认或加入：

- `research_period.start_date`：研究开始日期，格式 `YYYY-MM-DD`。
- `research_period.end_date`：研究结束日期，格式 `YYYY-MM-DD`。
- `article_content.enabled`：建议 `true`。
- `article_content.max_fetch_per_source_per_run`：建议先设为 `3`，避免过度访问媒体网站。
- `email.enabled`：不需要邮件提醒时设为 `false`。

`config.example.json` 只是字段结构示例，RSS 地址留空，不能直接作为正式监测配置。

### Secret B: `NEWS_MONITOR_STATE_KEY`

在本机 PowerShell 运行以下命令，复制输出字符串作为 Secret 值：

```powershell
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

请安全保存这个密钥。不要提交到仓库，也不要发给其他人。丢失后将无法解密之前的数据库快照。

## 3. 手动测试

进入仓库的 `Actions` 页面，选择 `News Monitor V6`，点击 `Run workflow`，选择 `main` 分支并启动。

首次运行没有历史数据库快照，会建立数据库并以 baseline 模式收集；之后每次运行会下载、解密上一次快照，追加新闻，再加密保存新快照。

## 4. 输出在哪里

每次成功运行会生成两个 Actions Artifacts：

- `news-monitor-state`：加密后的 SQLite 数据库，供下一次运行恢复。
- `coding-manifest`：仅含新闻元数据及空白人工编码字段，不包含 RSS 摘要和提取正文。

在 Actions 某次成功运行页面底部的 Artifacts 区域下载 `coding-manifest`，可供组员分配人工编码。完整数据库是加密状态文件，不能直接当 CSV 打开。

## 5. Public 仓库安全注意

- 不要将真实 `config.json` 上传到 Public 仓库。
- 不要提交邮箱密码、SMTP 凭证、状态加密密钥或 SQLite 数据库。
- 如果此前曾把凭证提交到公开仓库，删除文件并不能抹掉 Git 历史；应立即更换已经暴露的密码/密钥。
- 公开仓库中的 workflow、Python 代码和 `config.example.json` 是公开的。GitHub Actions Artifacts 的可见性受仓库权限影响，因此数据库快照必须先加密。
- 正文提取只访问公开可访问页面，不绕过登录、付费墙或反爬虫限制。提取失败时程序会保留状态供人工检查。

## 6. 时间与执行频率

定时表达式 `17 * * * *` 表示每小时 UTC 第 17 分钟运行一次。GitHub 定时任务可能延迟，不保证准点执行。手动运行可用于立即测试。
