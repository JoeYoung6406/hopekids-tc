#!/usr/bin/env python3
"""每月第三個週六：把「下個月」的教材從教材連結表同步到首頁 index.html。

更新兩個區塊（最新月份放最上層，前兩個月各自收合，更早的刪除）：
  - 本月投影片PPT：每週小班／大班（或合班）的教材 PPT
  - 本月課程範例：課程講義 + 故事圖

資料來源：教材連結表，每月一個分頁（例如「十月份」），需維持「知道連結的人可檢視」
  https://docs.google.com/spreadsheets/d/1c-AOXwkfi-nXFNZnsHSaM35Azlshv9yUOKMYmoXkp1Q
  欄位：大／小班、日期（例：第二週 10/11）、聖經、教材 PPT、課程講義、故事圖、勞作素材
  同一格「教材 PPT」有兩行連結時，依序為小班、大班（已用九月份分頁與實際連結核對）。

index.html 裡由本腳本管理的範圍：
  <!-- auto:slides ... --> … <!-- /auto:slides -->
  <!-- auto:course ... --> … <!-- /auto:course -->
每個月份包在 <!-- month:YYYY-MM --> … <!-- /month:YYYY-MM --> 之間。
腳本產生的連結帶 data-auto；重跑同一個月只會換掉 data-auto 的連結，手動加的（如故事教案頁）會保留。

用法：
  python .github/scripts/monthly_update.py                        # 上架下個月（台北時間）
  python .github/scripts/monthly_update.py --month 2026-10        # 指定月份
  python .github/scripts/monthly_update.py --third-saturday-only  # 排程用：不是第三個週六就直接結束
  python .github/scripts/monthly_update.py --dry-run              # 只印出差異，不寫檔
"""
import argparse
import csv
import datetime as dt
import difflib
import html
import io
import os
import re
import sys
import textwrap
import urllib.request
from pathlib import Path

SHEET_ID = "1c-AOXwkfi-nXFNZnsHSaM35Azlshv9yUOKMYmoXkp1Q"
INDEX = Path(__file__).resolve().parents[2] / "index.html"
KEEP_MONTHS = 3                      # 新月份 + 前兩個月
PPT_ORDER = ["小班", "大班"]          # 同一格兩行連結的順序
TAIPEI = dt.timezone(dt.timedelta(hours=8))

REGIONS = {
    "slides": {"tile": "📽️", "fold": "{m}月投影片PPT"},
    "course": {"tile": "🖼️", "fold": "{m}月課程範例"},
}

CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
          "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12}
TAB_RE = re.compile(r"^\s*(?:\d{4}\s*年?)?\s*(十[一二]?|[一二三四五六七八九]|\d{1,2})\s*月份?\s*$")
URL_RE = re.compile(r"https?://[^\s\"'<>]+")
DATE_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})")


def warn(msg):
    # GitHub Actions 會把 ::warning:: 顯示在執行摘要上
    print(f"::warning::{msg}" if os.environ.get("GITHUB_ACTIONS") else f"警告：{msg}")


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (hopekids-monthly-update)"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


# ---------- 讀教材連結表 ----------

def tab_month(name):
    m = TAB_RE.match(name)
    if not m:
        return None
    t = m.group(1)
    return int(t) if t.isdigit() else CN_NUM[t]


def list_tabs():
    page = fetch(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/htmlview")
    tabs = re.findall(r'\{name: "([^"]*)"[^}]*?gid: "(\d+)"', page)
    if not tabs:
        sys.exit("讀不到教材連結表的分頁清單（試算表是否仍為「知道連結的人可檢視」？）")
    return tabs


def find_col(header, *keywords):
    for i, name in enumerate(header):
        if any(k in name for k in keywords):
            return i
    sys.exit(f"教材連結表找不到欄位：{'/'.join(keywords)}（表頭：{header}）")


def load_weeks(year, month):
    """回傳該月每個主日的資料：date, scripture, classes[(班別, [PPT連結])], lesson[], story[]"""
    tabs = [(n, gid) for n, gid in list_tabs() if tab_month(n) == month]
    if not tabs:
        sys.exit(f"教材連結表沒有「{month}月份」分頁，請先建好分頁再手動執行 workflow。")

    for tab_name, gid in tabs:
        text = fetch(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}")
        rows = list(csv.reader(io.StringIO(text)))
        if not rows:
            continue
        header = [h.strip() for h in rows[0]]
        c_cls = find_col(header, "班")
        c_date = find_col(header, "日期")
        c_bible = find_col(header, "聖經", "經文")
        c_ppt = find_col(header, "PPT", "投影片")
        c_lesson = find_col(header, "講義")
        c_story = find_col(header, "故事圖")

        weeks, cur = [], None
        for row in rows[1:]:
            row += [""] * (len(header) - len(row))
            cell = lambda i: row[i].strip()
            d = DATE_RE.search(cell(c_date))
            if d:
                cur = {"date": dt.date(year, int(d.group(1)), int(d.group(2))),
                       "scripture": cell(c_bible), "classes": [], "lesson": [], "story": []}
                weeks.append(cur)
            if cur is None or not any(cell(i) for i in range(len(header))):
                continue
            cur["classes"].append((cell(c_cls), URL_RE.findall(cell(c_ppt))))
            cur["lesson"] += URL_RE.findall(cell(c_lesson))
            cur["story"] += URL_RE.findall(cell(c_story))

        weeks = [w for w in weeks if w["date"].month == month]
        if weeks:
            print(f"讀取分頁「{tab_name}」：{len(weeks)} 個主日")
            return weeks

    sys.exit(f"「{month}月份」分頁裡還沒有 {month} 月的日期與連結，請先填好再手動執行 workflow。")


# ---------- 產生連結卡 ----------

def link_html(label, url, date=None):
    lines = [f'<a class="link" data-auto href="{html.escape(url)}" target="_blank" rel="noopener">',
             f'  <span class="meta"><b>{html.escape(label)}</b></span>']
    if date:
        lines.append(f'  <span class="date">{date.month}/{date.day}</span>')
    lines += ['  <span class="arw">↗</span>', "</a>"]
    return "\n".join(lines)


def ppt_items(weeks):
    items = []
    for w in weeks:
        rows = [(cls, links) for cls, links in w["classes"] if links]
        names = {cls for cls, _ in w["classes"]}
        labeled = []
        if len(rows) == 1:
            cls, links = rows[0]
            if len(links) == 1:
                if "合" in cls or not {"大班", "小班"} <= names:
                    labeled = [(cls or "合班", links[0])]
                else:
                    labeled = [("大小班", links[0])]
            elif len(links) == len(PPT_ORDER):
                labeled = list(zip(PPT_ORDER, links))
            else:
                labeled = [(f"PPT {i}", u) for i, u in enumerate(links, 1)]
        else:
            for cls, links in rows:
                labeled += [(cls if len(links) == 1 else f"{cls} {i}", u) for i, u in enumerate(links, 1)]
        if not labeled:
            warn(f"{w['date']:%m/%d} 還沒有教材 PPT 連結")
        items += [link_html(f"{w['date']:%m%d} {name}", url) for name, url in labeled]
    return items


def link_key(url):
    """同一份 Google 文件（不同分頁）或同一個雲端資料夾視為同一個連結"""
    m = re.search(r"docs\.google\.com/document/d/([\w-]+)", url)
    if m:
        return m.group(1), f"https://docs.google.com/document/d/{m.group(1)}/edit"
    m = re.search(r"drive\.google\.com/drive/(?:u/\d+/)?folders/([\w-]+)", url)
    if m:
        return m.group(1), f"https://drive.google.com/drive/folders/{m.group(1)}"
    return url, url


def monthly_or_weekly(weeks, field, month, monthly_label, weekly_label):
    """全月共用一個連結 → 一張「N月xxx」；每週不同 → 每週一張並標日期"""
    per_week = [(w, w[field][0]) for w in weeks if w[field]]
    if not per_week:
        warn(f"{month} 月還沒有{weekly_label}連結")
        return []
    keys = {link_key(u)[0] for _, u in per_week}
    if len(keys) == 1:
        return [link_html(monthly_label.format(m=month), link_key(per_week[0][1])[1])]
    return [link_html(f"{weekly_label}｜{w['scripture']}" if w["scripture"] else weekly_label, u, w["date"])
            for w, u in per_week]


def course_items(weeks, month):
    return (monthly_or_weekly(weeks, "lesson", month, "{m}月課程範例", "課程範例")
            + monthly_or_weekly(weeks, "story", month, "{m}月故事圖", "故事圖"))


# ---------- 讀寫 index.html 的自動區塊 ----------

def region_re(name):
    return re.compile(rf"(^[ \t]*<!-- auto:{name}\b[^\n]*-->\n)(.*?)(^[ \t]*<!-- /auto:{name} -->)", re.S | re.M)


MONTH_RE = re.compile(r"^[ \t]*<!-- month:(\d{4}-\d{2}) -->\n(.*?)^[ \t]*<!-- /month:\1 -->", re.S | re.M)
CHUNK_RE = re.compile(r"<!--.*?-->|<(a|div)\b.*?</\1>", re.S)


def parse_months(body):
    return {key: textwrap.dedent(inner).strip("\n") for key, inner in MONTH_RE.findall(body)}


def merge_month(old_body, new_items):
    """換掉 data-auto 的連結，保留手動加的內容"""
    kept = []
    if old_body:
        kept = [m.group(0) for m in CHUNK_RE.finditer(old_body) if "data-auto" not in m.group(0)]
    return "\n".join(new_items + kept)


def render_block(key, body, indent):
    pad = " " * indent
    return "\n".join([f"{pad}<!-- month:{key} -->",
                      textwrap.indent(body, pad),
                      f"{pad}<!-- /month:{key} -->"])


def render_region(name, months):
    cfg = REGIONS[name]
    keys = sorted(months, reverse=True)[:KEEP_MONTHS]
    out = []
    for i, key in enumerate(keys):
        if i == 0:
            out.append(render_block(key, months[key], 4))
            continue
        out += ["",
                '      <details class="fold">',
                "        <summary>",
                f'          <span class="tile">{cfg["tile"]}</span>',
                f'          <span>{cfg["fold"].format(m=int(key[5:]))}</span>',
                '          <span class="caret">▼</span>',
                "        </summary>",
                "",
                render_block(key, months[key], 6),
                "      </details>"]
    return "\n".join(out) + "\n"


def update_region(page, name, key, items):
    m = region_re(name).search(page)
    if not m:
        sys.exit(f"index.html 找不到 <!-- auto:{name} --> 區塊標記")
    months = parse_months(m.group(2))
    months[key] = merge_month(months.get(key), items)
    return page[:m.start(2)] + render_region(name, months) + page[m.end(2):]


# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--month", help="要上架的月份 YYYY-MM（預設：台北時間的下個月）")
    ap.add_argument("--third-saturday-only", action="store_true", help="不是第三個週六就直接結束（排程用）")
    ap.add_argument("--dry-run", action="store_true", help="只印出差異，不寫檔")
    args = ap.parse_args()

    today = dt.datetime.now(TAIPEI).date()
    if args.third_saturday_only and not (today.weekday() == 5 and 15 <= today.day <= 21):
        print(f"今天（台北 {today}）不是當月第三個週六，略過。")
        return

    if args.month:
        if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", args.month):
            sys.exit(f"--month 格式要是 YYYY-MM：{args.month!r}")
        year, month = map(int, args.month.split("-"))
    else:
        year, month = (today.year + 1, 1) if today.month == 12 else (today.year, today.month + 1)
    key = f"{year}-{month:02d}"
    print(f"上架月份：{key}")

    weeks = load_weeks(year, month)
    slides, course = ppt_items(weeks), course_items(weeks, month)
    if not slides and not course:
        sys.exit(f"{key} 還沒有任何教材連結，先不更新。")

    with open(INDEX, encoding="utf-8", newline="") as f:
        raw = f.read()
    crlf = "\r\n" in raw
    page = raw.replace("\r\n", "\n")
    new = update_region(update_region(page, "slides", key, slides), "course", key, course)

    if new == page:
        print("index.html 已是最新，沒有變更。")
        return
    sys.stdout.writelines(difflib.unified_diff(page.splitlines(True), new.splitlines(True),
                                               "index.html", "index.html (new)"))
    if args.dry_run:
        print("\n（dry-run：未寫入）")
        return
    with open(INDEX, "w", encoding="utf-8", newline="") as f:
        f.write(new.replace("\n", "\r\n") if crlf else new)
    print(f"\n已更新 index.html：投影片 {len(slides)} 張、課程範例 {len(course)} 張")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    main()
