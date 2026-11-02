# News Monitor V5

用于香港/台湾新闻业研究的 RSS 新闻监测程序。

## 本版媒体
- 香港：香港電台 RTHK、香港標準 The Standard
- 台湾：中央社 CNA、自由時報 LTN

本版不再直接爬取新闻网页。优先使用公开 RSS；如果某媒体的全部主 RSS 都失败，才尝试该媒体的备用 RSS。所有主 RSS 都会读取，不再只读取第一个。

## 第一次运行
在 CMD：

```bat
cd C:\Users\Lenovo\Desktop\news_monitor_v5
python -m pip install -r requirements.txt
python news_monitor.py --once --baseline
```

然后：

```bat
python news_monitor.py --once
```

## 结果怎么看
程序会显示：

```text
香港電台 RTHK: feeds_ok=4, entries=..., new=..., strong_keyword_hits=...
```

- `feeds_ok`：成功读取的 RSS 数量
- `entries`：本次从 RSS 读取的条目数量（去重前）
- `new`：数据库中新加入的新闻数量
- `strong_keyword_hits`：命中强关键词的新增新闻数量

所以 `new=0` 不等于抓取失败。只要 `feeds_ok>0`，就代表 RSS 成功；可能只是新闻已经存在数据库。

## 重要
第一次运行建议用 `--baseline`。它会把当前新闻建立成基线，不发送提醒。以后正常运行即可。

程序只保存 RSS 提供的标题、摘要、链接、时间等字段；不会为了研究而强行抓取被媒体限制访问的全文。

## Excel
双击 `export.bat`，生成 `news_monitor_export.xlsx`。
