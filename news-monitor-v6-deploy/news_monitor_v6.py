#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""News Monitor V6

RSS-based news collection for the Hong Kong / Taiwan comparative study.

Main improvements over V5:
- Deduplicates by (media, URL), not URL alone.
- Migrates an existing V5 SQLite database without discarding collected rows.
- Supports an optional research date window; undated stories are retained and flagged.
- Optionally extracts publicly accessible article text; RSS summaries are never labelled as full text.
- Retries pending/failed text extraction at a controlled interval.
- Exports Excel-friendly CSV files and a blank manual-coding template.
- Supports one-shot execution for GitHub Actions / scheduled cloud jobs.

No paywall, login, or anti-bot bypass is attempted. Respect each publisher's terms and robots policy.
"""

import argparse
import csv
import json
import logging
import os
import re
import smtplib
import sqlite3
import time
from datetime import date, datetime, timedelta, timezone
from email.header import Header
from email.mime.text import MIMEText
from pathlib import Path
from urllib.parse import parse_qsl, quote_plus, urlencode, urlsplit, urlunsplit

import feedparser
import requests
from dateutil import parser as dateparser

BASE = Path(__file__).resolve().parent
CONFIG_PATH = BASE / "config.json"
DEFAULT_DB_PATH = BASE / "news_monitor.db"
DEFAULT_EXPORT_DIR = BASE / "exports"
LOG_PATH = BASE / "monitor.log"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "Chrome/129 Safari/537.36 NewsResearchMonitor/6.0"
    ),
    "Accept": "application/rss+xml, application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.5",
    "Accept-Language": "zh-TW,zh-HK,zh;q=0.9,en;q=0.6",
    "Cache-Control": "no-cache",
}

KEYWORDS = {
    "traffic": {
        "strong": [
            "流量變現", "流量变现", "點擊率", "点击率", "瀏覽量", "浏览量",
            "閱讀量", "阅读量", "流量收入", "熱搜榜", "热搜榜", "熱門搜尋",
            "热门搜索", "點閱率", "点阅率", "點閱量", "点阅量",
        ],
        "weak": [
            "流量", "點擊", "点击", "熱搜", "热搜", "熱門", "热门", "瘋傳",
            "疯传", "網友", "网友", "社交媒體", "社交媒体", "Facebook",
            "Instagram", "YouTube", "TikTok", "LINE", "轉發", "转发", "分享",
        ],
    },
    "commercial": {
        "strong": [
            "廣告收入", "广告收入", "訂閱收入", "订阅收入", "會員收入", "会员收入",
            "付費訂閱", "付费订阅", "營收下滑", "营收下滑", "商業模式", "商业模式",
            "流量變現", "流量变现", "裁員", "裁员", "虧損", "亏损", "廣告業務", "广告业务",
        ],
        "weak": [
            "廣告", "广告", "營收", "营收", "訂閱", "订阅", "會員", "会员",
            "贊助", "赞助", "商業", "商业", "成本", "付費", "付费",
        ],
    },
    "professionalism": {
        "strong": [
            "新聞倫理", "新闻伦理", "事實核查", "事实核查", "查證", "查证",
            "消息來源", "消息来源", "更正聲明", "更正声明", "澄清聲明", "澄清声明",
            "新聞自由", "新闻自由", "新聞專業主義", "新闻专业主义", "編採方針",
            "编采方针", "查核", "事实查核",
        ],
        "weak": ["更正", "澄清", "採訪", "采访", "編採", "编采", "記者", "记者", "消息人士"],
    },
}


def configure_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8"), logging.StreamHandler()],
        force=True,
    )


def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"找不到配置文件：{CONFIG_PATH}\n请将原项目的 config.json 放在本程序同一文件夹。"
        )
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    if not isinstance(cfg.get("sources"), dict) or not cfg["sources"]:
        raise ValueError("config.json 必须包含非空的 sources 对象。")
    return cfg


def resolve_path(value, default):
    raw = Path(value or default)
    return raw if raw.is_absolute() else BASE / raw


def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def clean_article_text(value):
    """Normalize article text while retaining paragraph boundaries."""
    lines = [re.sub(r"[\t\xa0 ]+", " ", line).strip() for line in str(value or "").splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def parse_time(entry):
    for key in ("published", "updated", "created"):
        value = entry.get(key)
        if value:
            try:
                return dateparser.parse(value, fuzzy=True).isoformat()
            except Exception:
                pass
    for key in ("published_parsed", "updated_parsed"):
        value = entry.get(key)
        if value:
            try:
                return datetime(*value[:6], tzinfo=timezone.utc).isoformat()
            except Exception:
                pass
    return ""


def parse_date_arg(value, label):
    if not value:
        return None
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise ValueError(f"{label} 必须使用 YYYY-MM-DD 格式，当前值：{value}") from exc


def classify_date(published_time, start_date=None, end_date=None):
    """Return (include, status). Unknown dates are kept for manual review."""
    if not published_time:
        return True, "unknown"
    try:
        published_date = dateparser.parse(published_time, fuzzy=True).date()
    except Exception:
        return True, "unknown"
    if not start_date and not end_date:
        return True, "not_filtered"
    if start_date and published_date < start_date:
        return False, "before_period"
    if end_date and published_date > end_date:
        return False, "after_period"
    return True, "in_period"


def score_keywords(text):
    low = (text or "").lower()
    all_hits, groups, research_hits, scores = [], [], [], {}
    for group, cfg in KEYWORDS.items():
        strong = sorted({w for w in cfg["strong"] if w.lower() in low})
        weak = sorted({w for w in cfg["weak"] if w.lower() in low})
        score = min(5, len(strong) * 2 + len(weak))
        scores[group] = score
        if score:
            groups.append(group)
        all_hits.extend(strong + weak)
        research_hits.extend(f"{group}:强:{word}" for word in strong)
    return sorted(set(all_hits)), groups, scores, research_hits


def normalize_url(url):
    """Remove fragments and common tracking parameters, preserving meaningful query parameters."""
    raw = clean(url)
    if not raw:
        return ""
    try:
        parts = urlsplit(raw)
        tracking = ("utm_", "fbclid", "gclid", "mc_cid", "mc_eid", "ref_src")
        query = [
            (key, value)
            for key, value in parse_qsl(parts.query, keep_blank_values=True)
            if not key.lower().startswith("utm_") and key.lower() not in tracking
        ]
        return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query, doseq=True), ""))
    except Exception:
        return raw.split("#", 1)[0]


def infer_region(media_name, source_cfg):
    explicit = clean(source_cfg.get("region", ""))
    if explicit:
        return explicit
    name = str(media_name).lower()
    if any(x in name for x in ("rthk", "standard", "香港", "hong kong")):
        return "Hong Kong"
    if any(x in name for x in ("cna", "ltn", "自由時報", "自由时报", "liberty", "taiwan", "台灣", "台湾")):
        return "Taiwan"
    return ""


def make_google_news_rss(domain, hl="zh-TW", gl="TW", ceid="TW:zh-Hant"):
    q = quote_plus(f"site:{domain}")
    return f"https://news.google.com/rss/search?q={q}&hl={hl}&gl={gl}&ceid={quote_plus(ceid)}"


def fetch_feed(session, url, timeout):
    last_error = ""
    for attempt in range(3):
        try:
            response = session.get(url, timeout=timeout, headers=HEADERS, allow_redirects=True)
            content_type = response.headers.get("content-type", "")
            if response.status_code == 200 and len(response.content) > 100:
                parsed = feedparser.parse(response.content)
                if parsed.entries:
                    return parsed, f"ok ({len(parsed.entries)} entries; {content_type or 'unknown'})"
                if not parsed.bozo:
                    return parsed, f"ok (0 entries; {content_type or 'unknown'})"
                last_error = f"parse warning: {parsed.bozo_exception}"
            else:
                last_error = f"HTTP {response.status_code} ({content_type or 'unknown'})"
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        if attempt < 2:
            time.sleep(1.5 * (attempt + 1))
    return None, last_error


def fetch_article_content(session, url, timeout, max_chars=30000):
    """Extract visible article text from a public page; never bypass access controls."""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        return "", "dependency_missing", "未安装 beautifulsoup4"

    try:
        response = session.get(url, timeout=timeout, headers=HEADERS, allow_redirects=True)
        if response.status_code != 200:
            return "", "fetch_failed", f"HTTP {response.status_code}"
        content_type = response.headers.get("content-type", "").lower()
        if response.encoding and response.encoding.lower() in ("iso-8859-1", "ascii"):
            detected_encoding = response.apparent_encoding
            if detected_encoding:
                response.encoding = detected_encoding
        if content_type and not any(x in content_type for x in ("html", "xhtml", "text/")):
            return "", "not_html", f"非 HTML 内容：{content_type[:100]}"

        soup = BeautifulSoup(response.text, "html.parser")
        for node in soup(["script", "style", "noscript", "svg", "nav", "footer", "header", "aside", "form", "button"]):
            node.decompose()

        noise_pattern = re.compile(
            r"(comment|recommend|related|social|share|advert|cookie|breadcrumb|footer|navigation)",
            re.I,
        )
        for node in soup.find_all(attrs={"class": noise_pattern}):
            node.decompose()
        for node in soup.find_all(attrs={"id": noise_pattern}):
            node.decompose()

        candidates = []
        selectors = [
            "[itemprop='articleBody']", "article", ".article-body", ".article-content",
            ".story-body", ".story-content", ".post-content", ".entry-content", "main",
        ]
        for selector in selectors:
            for node in soup.select(selector):
                text = clean_article_text(node.get_text("\n", strip=True))
                if len(text) > 0:
                    candidates.append(text)
        if candidates:
            extracted = max(candidates, key=len)
        else:
            root = soup.body or soup
            paragraphs = [clean(p.get_text(" ", strip=True)) for p in root.find_all("p")]
            extracted = "\n".join(p for p in paragraphs if len(p) >= 25)

        extracted = clean_article_text(extracted)
        if len(extracted) > max_chars:
            extracted = extracted[:max_chars].rsplit("\n", 1)[0]
        if len(extracted) >= 500:
            return extracted, "full", ""
        if len(extracted) >= 120:
            return extracted, "partial", "正文提取长度较短，请人工检查"
        return "", "rss_summary", "页面未找到足够长的正文文本"
    except requests.RequestException as exc:
        return "", "fetch_failed", f"{type(exc).__name__}: {exc}"
    except Exception as exc:
        return "", "extract_failed", f"{type(exc).__name__}: {exc}"


CREATE_TABLE_SQL = """CREATE TABLE articles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    media TEXT NOT NULL,
    region TEXT DEFAULT '',
    title TEXT,
    url TEXT NOT NULL,
    published_time TEXT,
    date_status TEXT DEFAULT 'unknown',
    author TEXT,
    summary TEXT,
    content TEXT,
    content_status TEXT DEFAULT 'rss_summary',
    content_error TEXT DEFAULT '',
    content_checked_time TEXT,
    matched_keywords TEXT,
    keyword_groups TEXT,
    scraped_time TEXT NOT NULL,
    last_seen_time TEXT,
    traffic_score INTEGER DEFAULT 0,
    commercial_score INTEGER DEFAULT 0,
    professionalism_score INTEGER DEFAULT 0,
    research_hits TEXT DEFAULT '',
    source_feed TEXT,
    UNIQUE(media, url)
)"""


def table_exists(con, table_name):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table_name,)
    ).fetchone() is not None


def has_media_url_unique(con):
    for index in con.execute("PRAGMA index_list(articles)").fetchall():
        if int(index[2]) == 1:
            columns = [row[2] for row in con.execute(f"PRAGMA index_info('{index[1]}')").fetchall()]
            if columns == ["media", "url"]:
                return True
    return False


def init_db(db_path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path), timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=30000")
    con.execute("PRAGMA journal_mode=WAL")

    if not table_exists(con, "articles"):
        con.execute(CREATE_TABLE_SQL)
        logging.info("Created new V6 database: %s", db_path)
    elif not has_media_url_unique(con):
        # Rebuild the V5 table. This removes the old global UNIQUE(url) rule,
        # which could incorrectly suppress the same URL published by another outlet.
        old_columns = [row[1] for row in con.execute("PRAGMA table_info(articles)").fetchall()]
        old_rows = [dict(row) for row in con.execute("SELECT * FROM articles").fetchall()]
        con.execute("ALTER TABLE articles RENAME TO articles_legacy_v5")
        con.execute(CREATE_TABLE_SQL)
        new_columns = [
            "id", "media", "region", "title", "url", "published_time", "date_status",
            "author", "summary", "content", "content_status", "content_error",
            "content_checked_time", "matched_keywords", "keyword_groups", "scraped_time",
            "last_seen_time", "traffic_score", "commercial_score", "professionalism_score",
            "research_hits", "source_feed",
        ]
        migrated = 0
        for old in old_rows:
            row = {key: old.get(key) for key in old_columns}
            values = {
                "id": row.get("id"),
                "media": row.get("media") or "Unknown",
                "region": infer_region(row.get("media") or "", {}),
                "title": row.get("title") or "",
                "url": row.get("url") or "",
                "published_time": row.get("published_time") or "",
                "date_status": "unknown",
                "author": row.get("author") or "",
                "summary": row.get("summary") or "",
                "content": row.get("content") or row.get("summary") or "",
                "content_status": "rss_summary",
                "content_error": "",
                "content_checked_time": None,
                "matched_keywords": row.get("matched_keywords") or "[]",
                "keyword_groups": row.get("keyword_groups") or "[]",
                "scraped_time": row.get("scraped_time") or now_iso(),
                "last_seen_time": row.get("scraped_time") or now_iso(),
                "traffic_score": row.get("traffic_score") or 0,
                "commercial_score": row.get("commercial_score") or 0,
                "professionalism_score": row.get("professionalism_score") or 0,
                "research_hits": row.get("research_hits") or "[]",
                "source_feed": row.get("source_feed") or "",
            }
            placeholders = ",".join("?" for _ in new_columns)
            con.execute(
                f"INSERT OR IGNORE INTO articles ({','.join(new_columns)}) VALUES ({placeholders})",
                [values[column] for column in new_columns],
            )
            migrated += 1
        con.execute("DROP TABLE articles_legacy_v5")
        logging.warning(
            "Migrated %d existing rows to V6. V5 may already have omitted cross-media duplicate URLs; those historical rows cannot be reconstructed by migration.",
            migrated,
        )
    else:
        # Add forward-compatible columns if a partially upgraded database exists.
        existing = {row[1] for row in con.execute("PRAGMA table_info(articles)").fetchall()}
        additions = {
            "region": "TEXT DEFAULT ''",
            "date_status": "TEXT DEFAULT 'unknown'",
            "content_status": "TEXT DEFAULT 'rss_summary'",
            "content_error": "TEXT DEFAULT ''",
            "content_checked_time": "TEXT",
            "last_seen_time": "TEXT",
        }
        for column, definition in additions.items():
            if column not in existing:
                con.execute(f"ALTER TABLE articles ADD COLUMN {column} {definition}")

    con.execute("CREATE INDEX IF NOT EXISTS idx_media_time ON articles(media, published_time)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_region_time ON articles(region, published_time)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_content_status ON articles(content_status)")
    con.commit()
    return con


def should_retry_content(existing, retry_hours):
    if existing is None:
        return True
    status = existing["content_status"] or "rss_summary"
    if status == "full":
        return False
    checked = existing["content_checked_time"]
    if not checked:
        return True
    try:
        checked_at = dateparser.parse(checked)
        if checked_at.tzinfo is None:
            checked_at = checked_at.replace(tzinfo=timezone.utc)
        return datetime.now(timezone.utc) - checked_at >= timedelta(hours=retry_hours)
    except Exception:
        return True


def save_article(con, article):
    text_for_scoring = "\n".join(
        x for x in (article.get("title", ""), article.get("summary", ""), article.get("content", "")) if x
    )
    hits, groups, scores, research_hits = score_keywords(text_for_scoring)
    now = now_iso()
    existing = con.execute(
        "SELECT * FROM articles WHERE media=? AND url=?",
        (article["media"], article["url"]),
    ).fetchone()

    if existing:
        # Preserve an already extracted full article if a later RSS response only has a summary.
        incoming_content = article.get("content") or article.get("summary") or ""
        old_content = existing["content"] or ""
        incoming_status = article.get("content_status", "rss_summary")
        old_status = existing["content_status"] or "rss_summary"
        if old_status == "full":
            # Never replace an already extracted full article with a shorter retry.
            use_incoming = incoming_status == "full" and len(incoming_content) > len(old_content)
        elif incoming_status == "full":
            use_incoming = True
        elif incoming_status == "partial":
            use_incoming = len(incoming_content) >= len(old_content)
        else:
            use_incoming = old_status not in ("full",) and len(incoming_content) >= len(old_content)
        chosen_content = incoming_content if use_incoming else old_content
        chosen_status = incoming_status if use_incoming else old_status
        chosen_checked = article.get("content_checked_time") or existing["content_checked_time"]
        chosen_error = article.get("content_error", "") if chosen_status == incoming_status else existing["content_error"]
        con.execute(
            """UPDATE articles SET
                region=CASE WHEN ?<>'' THEN ? ELSE region END,
                title=CASE WHEN ?<>'' THEN ? ELSE title END,
                published_time=CASE WHEN ?<>'' THEN ? ELSE published_time END,
                date_status=?,
                author=CASE WHEN ?<>'' THEN ? ELSE author END,
                summary=CASE WHEN ?<>'' THEN ? ELSE summary END,
                content=?, content_status=?, content_error=?, content_checked_time=?,
                matched_keywords=?, keyword_groups=?, traffic_score=?, commercial_score=?,
                professionalism_score=?, research_hits=?, source_feed=?, last_seen_time=?
              WHERE media=? AND url=?""",
            (
                article.get("region", ""), article.get("region", ""),
                article.get("title", ""), article.get("title", ""),
                article.get("published_time", ""), article.get("published_time", ""),
                article.get("date_status", "unknown"),
                article.get("author", ""), article.get("author", ""),
                article.get("summary", ""), article.get("summary", ""),
                chosen_content, chosen_status, chosen_error, chosen_checked,
                json.dumps(hits, ensure_ascii=False), json.dumps(groups, ensure_ascii=False),
                scores["traffic"], scores["commercial"], scores["professionalism"],
                json.dumps(research_hits, ensure_ascii=False), article.get("source_feed", ""), now,
                article["media"], article["url"],
            ),
        )
        con.commit()
        return False, research_hits, scores

    con.execute(
        """INSERT INTO articles (
            media, region, title, url, published_time, date_status, author, summary, content,
            content_status, content_error, content_checked_time, matched_keywords, keyword_groups,
            scraped_time, last_seen_time, traffic_score, commercial_score,
            professionalism_score, research_hits, source_feed
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            article["media"], article.get("region", ""), article["title"], article["url"],
            article.get("published_time", ""), article.get("date_status", "unknown"),
            article.get("author", ""), article.get("summary", ""),
            article.get("content") or article.get("summary", ""),
            article.get("content_status", "rss_summary"), article.get("content_error", ""),
            article.get("content_checked_time"), json.dumps(hits, ensure_ascii=False),
            json.dumps(groups, ensure_ascii=False), now, now, scores["traffic"],
            scores["commercial"], scores["professionalism"],
            json.dumps(research_hits, ensure_ascii=False), article.get("source_feed", ""),
        ),
    )
    con.commit()
    return True, research_hits, scores


def send_email(cfg, hits):
    email_cfg = cfg.get("email", {})
    if not email_cfg.get("enabled") or not hits:
        return
    required = ("from", "to", "smtp_host", "username", "password")
    missing = [key for key in required if not email_cfg.get(key)]
    if missing:
        raise ValueError("邮件提醒配置缺少字段：" + ", ".join(missing))
    body = "\n\n".join(
        f"[{item['media']}] {item['title']}\n{item['url']}\n研究命中：{', '.join(item['research_hits'])}"
        for item in hits
    )
    message = MIMEText(body, "plain", "utf-8")
    message["Subject"] = Header(f"新闻监测提醒：{len(hits)} 条强关键词命中", "utf-8")
    message["From"] = email_cfg["from"]
    message["To"] = email_cfg["to"]
    with smtplib.SMTP_SSL(
        email_cfg["smtp_host"], int(email_cfg.get("smtp_port", 465)), timeout=20
    ) as smtp:
        smtp.login(email_cfg["username"], email_cfg["password"])
        smtp.send_message(message)


def process_entry(session, con, media, region, entry, feed_url, cfg, date_window, fetch_budget):
    url = normalize_url(entry.get("link", ""))
    title = clean(entry.get("title", ""))
    if not url or not title:
        return "skipped", None

    published_time = parse_time(entry)
    include, date_status = classify_date(published_time, *date_window)
    if not include:
        return "outside_period", None

    summary = clean(entry.get("summary", "") or entry.get("description", ""))
    author = clean(entry.get("author", ""))
    existing = con.execute(
        "SELECT * FROM articles WHERE media=? AND url=?", (media, url)
    ).fetchone()

    content_cfg = cfg.get("article_content", {})
    content_enabled = bool(content_cfg.get("enabled", True)) and not cfg.get("_no_content", False)
    if existing:
        if not content_enabled or not should_retry_content(existing, int(content_cfg.get("retry_after_hours", 24))):
            con.execute("UPDATE articles SET last_seen_time=? WHERE media=? AND url=?", (now_iso(), media, url))
            con.commit()
            return "existing", None

    content = summary
    content_status = "disabled" if not content_enabled else "rss_summary"
    content_error = ""
    content_checked_time = now_iso() if not content_enabled else None

    per_source_limit = int(content_cfg.get("max_fetch_per_source_per_run", 10))
    if content_enabled and fetch_budget.get(media, 0) < per_source_limit:
        fetch_budget[media] = fetch_budget.get(media, 0) + 1
        timeout = int(cfg.get("request_timeout", 20))
        max_chars = int(content_cfg.get("max_chars", 30000))
        extracted, status, error = fetch_article_content(session, url, timeout, max_chars)
        content_status = status
        content_error = error
        content_checked_time = now_iso()
        if extracted:
            content = extracted
        delay = float(content_cfg.get("request_delay_seconds", 1.0))
        if delay > 0:
            time.sleep(delay)
    elif content_enabled:
        content_status = "pending"
        content_error = "本次运行达到正文提取数量上限，等待下次运行"

    article = {
        "media": media,
        "region": region,
        "title": title,
        "url": url,
        "published_time": published_time,
        "date_status": date_status,
        "author": author,
        "summary": summary,
        "content": content,
        "content_status": content_status,
        "content_error": content_error,
        "content_checked_time": content_checked_time,
        "source_feed": feed_url,
    }
    inserted, research_hits, _scores = save_article(con, article)
    return ("new" if inserted else "updated"), {
        "media": media,
        "title": title,
        "url": url,
        "research_hits": research_hits,
    }


def process_source(session, con, name, source_cfg, global_cfg, baseline, date_window, fetch_budget):
    feeds = list(source_cfg.get("feeds", []))
    fallback = list(source_cfg.get("fallback_feeds", []))
    max_entries = int(source_cfg.get("max_entries", 100))
    timeout = int(global_cfg.get("request_timeout", 20))
    region = infer_region(name, source_cfg)
    seen_urls = set()
    counters = {key: 0 for key in (
        "new", "updated", "existing", "outside_period", "entries", "ok_feeds", "failed_feeds", "content_hits"
    )}
    alerts = []

    logging.info("\n[%s] checking RSS (%s)", name, region or "region not configured")

    def process_feed_list(feed_list, label):
        for feed_url in feed_list:
            logging.info("  %s: %s", label, feed_url)
            feed, status = fetch_feed(session, feed_url, timeout)
            logging.info("    %s", status)
            if feed is None:
                counters["failed_feeds"] += 1
                continue
            counters["ok_feeds"] += 1
            entries = feed.entries[:max_entries]
            counters["entries"] += len(entries)
            for entry in entries:
                url = normalize_url(entry.get("link", ""))
                if not url or url in seen_urls:
                    continue
                seen_urls.add(url)
                result, alert = process_entry(
                    session, con, name, region, entry, feed_url, global_cfg, date_window, fetch_budget
                )
                if result in counters:
                    counters[result] += 1
                if alert and result in ("new", "updated") and alert["research_hits"] and not baseline:
                    counters["content_hits"] += 1
                    alerts.append(alert)

    process_feed_list(feeds, "primary")
    if counters["ok_feeds"] == 0 and fallback:
        logging.info("  all primary feeds failed; trying fallback feeds")
        process_feed_list(fallback, "fallback")

    logging.info(
        "  result: feeds_ok=%d, failed_feeds=%d, entries=%d, new=%d, updated=%d, existing=%d, outside_period=%d, strong_keyword_hits=%d",
        counters["ok_feeds"], counters["failed_feeds"], counters["entries"], counters["new"],
        counters["updated"], counters["existing"], counters["outside_period"], counters["content_hits"],
    )
    if counters["ok_feeds"] == 0:
        logging.error("[%s] all primary/fallback RSS feeds failed", name)
    return counters, alerts


def get_date_window(cfg, cli_start=None, cli_end=None):
    period = cfg.get("research_period", {})
    start = parse_date_arg(cli_start or period.get("start_date") or cfg.get("start_date"), "start_date")
    end = parse_date_arg(cli_end or period.get("end_date") or cfg.get("end_date"), "end_date")
    if start and end and start > end:
        raise ValueError("研究开始日期不能晚于结束日期。")
    return start, end


def csv_safe(value):
    if value is None:
        return ""
    return value


def write_csv(path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({key: csv_safe(row.get(key, "")) for key in fieldnames})


def export_data(con, export_dir, date_window=None):
    export_dir.mkdir(parents=True, exist_ok=True)
    all_rows = [dict(row) for row in con.execute(
        "SELECT * FROM articles ORDER BY published_time, media, title"
    ).fetchall()]
    rows = all_rows
    outside_rows = []
    if date_window and (date_window[0] or date_window[1]):
        rows = []
        for row in all_rows:
            include, status = classify_date(row.get("published_time", ""), *date_window)
            row["date_status"] = status
            (rows if include else outside_rows).append(row)
        write_csv(export_dir / "outside_period.csv", outside_rows, [
            "id", "media", "region", "title", "url", "published_time", "date_status", "summary"
        ])
        # Preserve a complete archive separately, while the research CSVs contain only the study window.
        if outside_rows:
            write_csv(export_dir / "database_archive.csv", all_rows, [
                "id", "media", "region", "title", "url", "published_time", "date_status", "summary",
                "content", "content_status", "scraped_time", "source_feed"
            ])
    if not rows:
        logging.warning("研究日期范围内目前没有新闻记录，仍会生成空 CSV 和编码模板。")

    fields = [
        "id", "media", "region", "title", "url", "published_time", "date_status", "author",
        "summary", "content", "content_status", "content_error", "content_checked_time",
        "matched_keywords", "keyword_groups", "traffic_score", "commercial_score",
        "professionalism_score", "research_hits", "scraped_time", "last_seen_time", "source_feed",
    ]
    write_csv(export_dir / "all_articles.csv", rows, fields)

    by_media = {}
    for row in rows:
        by_media.setdefault(row.get("media") or "Unknown", []).append(row)
    for media, media_rows in by_media.items():
        slug = re.sub(r"[^A-Za-z0-9_-]+", "_", media).strip("_") or "media"
        write_csv(export_dir / f"media_{slug}.csv", media_rows, fields)

    for region, filename in (("Hong Kong", "hong_kong.csv"), ("Taiwan", "taiwan.csv")):
        region_rows = [row for row in rows if row.get("region") == region]
        write_csv(export_dir / filename, region_rows, fields)

    hit_rows = [row for row in rows if row.get("keyword_groups") not in (None, "", "[]")]
    write_csv(export_dir / "keyword_hits.csv", hit_rows, fields)

    coding_fields = [
        "id", "media", "region", "title", "url", "published_time", "date_status",
        "summary", "content", "content_status",
        "manual_news_category", "manual_headline_style", "clickbait_tendency_1_to_5",
        "social_media_elements_yes_no", "manual_reporting_style", "coder", "coding_notes",
    ]
    coding_rows = []
    for row in rows:
        coding_rows.append({**row, "manual_news_category": "", "manual_headline_style": "",
                            "clickbait_tendency_1_to_5": "", "social_media_elements_yes_no": "",
                            "manual_reporting_style": "", "coder": "", "coding_notes": ""})
    write_csv(export_dir / "coding_template.csv", coding_rows, coding_fields)

    logging.info("CSV exports updated: %s (%d articles)", export_dir, len(rows))
    return len(rows)


def run_once(cfg, baseline=False, db_path=None, export_dir=None, start_date=None, end_date=None):
    db_path = db_path or resolve_path(os.getenv("NEWS_MONITOR_DB") or cfg.get("database_path"), DEFAULT_DB_PATH)
    export_dir = export_dir or resolve_path(os.getenv("NEWS_MONITOR_EXPORT_DIR") or cfg.get("export_dir"), DEFAULT_EXPORT_DIR)
    date_window = get_date_window(cfg, start_date, end_date)
    logging.info("Research date window: %s to %s", date_window[0] or "not set", date_window[1] or "not set")

    con = init_db(db_path)
    # Backfill region labels for rows collected by V5 or by configs that did not define region.
    for media_name, source_cfg in cfg["sources"].items():
        region = infer_region(media_name, source_cfg)
        if region:
            con.execute(
                "UPDATE articles SET region=? WHERE media=? AND (region IS NULL OR region='')",
                (region, media_name),
            )
    con.commit()
    session = requests.Session()
    cfg["_alerts"] = []
    fetch_budget = {}
    totals = {}
    try:
        for name, source_cfg in cfg["sources"].items():
            counters, alerts = process_source(
                session, con, name, source_cfg, cfg, baseline, date_window, fetch_budget
            )
            totals[name] = counters
            cfg["_alerts"].extend(alerts)

        if cfg["_alerts"] and not baseline:
            try:
                send_email(cfg, cfg["_alerts"])
                logging.info("email sent: %d", len(cfg["_alerts"]))
            except Exception as exc:
                logging.error("email failed: %s", exc)

        total_rows = con.execute("SELECT COUNT(*) FROM articles").fetchone()[0]
        logging.info("\n=== Database summary: %d unique media-article records ===", total_rows)
        for media, counters in totals.items():
            logging.info("%s: %s", media, json.dumps(counters, ensure_ascii=False))
        export_data(con, export_dir, date_window)
        return totals
    finally:
        session.close()
        con.close()


def main():
    parser = argparse.ArgumentParser(description="RSS News Monitor V6")
    parser.add_argument("--once", action="store_true", help="执行一次后退出（适合 GitHub Actions）")
    parser.add_argument("--baseline", action="store_true", help="基线收集，不发送关键词邮件提醒")
    parser.add_argument("--export-only", action="store_true", help="只从现有数据库重新导出 CSV")
    parser.add_argument("--start-date", help="覆盖配置中的研究开始日期：YYYY-MM-DD")
    parser.add_argument("--end-date", help="覆盖配置中的研究结束日期：YYYY-MM-DD")
    parser.add_argument("--db", help="指定 SQLite 数据库路径")
    parser.add_argument("--export-dir", help="指定 CSV 输出目录")
    parser.add_argument("--no-content", action="store_true", help="本次只收集 RSS，不抓取新闻正文")
    args = parser.parse_args()

    cfg = load_config()
    if args.no_content:
        cfg["_no_content"] = True

    db_path = resolve_path(args.db or os.getenv("NEWS_MONITOR_DB") or cfg.get("database_path"), DEFAULT_DB_PATH)
    export_dir = resolve_path(args.export_dir or os.getenv("NEWS_MONITOR_EXPORT_DIR") or cfg.get("export_dir"), DEFAULT_EXPORT_DIR)

    if args.export_only:
        con = init_db(db_path)
        try:
            export_data(con, export_dir, get_date_window(cfg, args.start_date, args.end_date))
        finally:
            con.close()
        return

    if args.once:
        run_once(cfg, args.baseline, db_path, export_dir, args.start_date, args.end_date)
        return

    interval_minutes = max(1, int(cfg.get("interval_minutes", 10)))
    logging.info("News Monitor V6 started; interval=%d minutes", interval_minutes)
    while True:
        try:
            run_once(cfg, False, db_path, export_dir, args.start_date, args.end_date)
        except Exception:
            logging.exception("Scheduled collection cycle failed")
        time.sleep(interval_minutes * 60)


if __name__ == "__main__":
    configure_logging()
    main()
