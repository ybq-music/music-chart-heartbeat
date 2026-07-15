from __future__ import annotations

import csv
import html
import json
import math
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "outputs"
CONFIG_PATH = DATA_DIR / "daily_run_config.json"
HEARTBEAT_CSV = DATA_DIR / "daily_heartbeat.csv"
DAILY_STATUS_CSV = DATA_DIR / "daily_status.csv"
RUN_STATUS_MD = ROOT / "RUN_STATUS.md"

TZ = ZoneInfo("Asia/Shanghai")
TARGET_DAYS = int(os.getenv("TARGET_DAYS", "7"))
REQUEST_TIMEOUT_SECONDS = int(os.getenv("REQUEST_TIMEOUT_SECONDS", "30"))
COLLECTION_SCHEMA = "music-weibo-v1"
COLLECTION_START_ON = os.getenv("COLLECTION_START_ON", "").strip()


DETAIL_FIELDS = [
    "snapshot_date",
    "captured_at",
    "platform",
    "chart",
    "official_chart",
    "platform_date",
    "rank",
    "song_name",
    "artist",
    "chart_metric",
    "duration",
    "song_id",
    "album_id",
    "source_url",
]

HEARTBEAT_FIELDS = [
    "run_id",
    "snapshot_date",
    "status",
    "platform",
    "chart",
    "official_chart",
    "row_count",
    "started_at",
    "finished_at",
    "error_message",
]

DAILY_STATUS_FIELDS = [
    "snapshot_date",
    "status",
    "row_count",
    "xlsx_path",
    "completed_at",
    "message",
]


@dataclass(frozen=True)
class ChartSource:
    platform: str
    chart: str
    official_chart: str
    kind: str
    url: str
    rank_id: str
    expected_rows: int
    note: str
    minimum_rows: int = 0
    variable_rows: bool = False


CHARTS = [
    ChartSource(
        platform="QQ音乐",
        chart="热歌榜",
        official_chart="热歌榜",
        kind="qq",
        rank_id="26",
        expected_rows=100,
        url="https://y.qq.com/n/ryqq/toplist/26",
        note="QQ音乐热歌榜，取公开榜单接口前100名。",
    ),
    ChartSource(
        platform="QQ音乐",
        chart="飙升榜",
        official_chart="飙升榜",
        kind="qq",
        rank_id="62",
        expected_rows=100,
        url="https://y.qq.com/n/ryqq/toplist/62",
        note="QQ音乐飙升榜，取公开榜单接口前100名。",
    ),
    ChartSource(
        platform="酷狗音乐",
        chart="热歌榜",
        official_chart="酷狗TOP500",
        kind="kugou",
        rank_id="8888",
        expected_rows=500,
        url="https://www.kugou.com/yy/rank/home/1-8888.html?from=rank",
        note="酷狗没有直接命名为热歌榜的入口，本任务采用官方酷狗TOP500作为热歌榜口径。",
    ),
    ChartSource(
        platform="酷狗音乐",
        chart="飙升榜",
        official_chart="酷狗飙升榜",
        kind="kugou",
        rank_id="6666",
        expected_rows=100,
        url="https://www.kugou.com/yy/rank/home/1-6666.html?from=rank",
        note="酷狗飙升榜，按公开榜单分页抓取。",
    ),
    ChartSource(
        platform="微博",
        chart="热搜榜",
        official_chart="微博热搜榜",
        kind="weibo_hot",
        rank_id="",
        expected_rows=50,
        url="https://weibo.com/ajax/statuses/hot_band",
        note="微博公开热搜JSON，剔除广告位后取前50名。",
    ),
    ChartSource(
        platform="微博",
        chart="文娱热搜榜",
        official_chart="微博热搜榜-文娱条目",
        kind="weibo_entertainment",
        rank_id="",
        expected_rows=50,
        minimum_rows=1,
        variable_rows=True,
        url="https://weibo.com/ajax/statuses/hot_band",
        note="从微博热搜JSON中提取Entertainment/艺人/剧集/电影/综艺/音乐等文娱条目，数量随当日榜单浮动。",
    ),
]


def now_local() -> datetime:
    return datetime.now(TZ)


def iso(dt: datetime) -> str:
    return dt.replace(microsecond=0).isoformat()


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def load_or_create_config(today: date) -> dict:
    ensure_dirs()
    if CONFIG_PATH.exists():
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        changed = False
        configured_start = COLLECTION_START_ON or str(config.get("collection_start_on", "")).strip()
        if not configured_start:
            configured_start = today.isoformat()
        updates = {
            "collection_schema": COLLECTION_SCHEMA,
            "collection_start_on": configured_start,
            "target_days": TARGET_DAYS,
        }
        for key, value in updates.items():
            if config.get(key) != value:
                config[key] = value
                changed = True
        if changed:
            CONFIG_PATH.write_text(
                json.dumps(config, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        return config

    config = {
        "timezone": "Asia/Shanghai",
        "target_days": TARGET_DAYS,
        "started_on": today.isoformat(),
        "collection_schema": COLLECTION_SCHEMA,
        "collection_start_on": COLLECTION_START_ON or today.isoformat(),
        "created_at": iso(now_local()),
    }
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return config


def request_referer(url: str) -> str:
    if "weibo.com" in url:
        return "https://weibo.com/"
    if "kugou.com" in url:
        return "https://www.kugou.com/"
    return "https://y.qq.com/"


def fetch_text(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json,text/plain,text/html,*/*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Referer": request_referer(url),
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
            ),
        },
    )
    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                return response.read().decode(charset, errors="replace")
        except Exception as exc:
            last_error = exc
            if attempt < 3:
                time.sleep(attempt * 2)
    raise RuntimeError(f"请求失败：{url}；原因：{last_error}")


def read_csv(path: Path, fields: list[str]) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        return [{field: row.get(field, "") for field in fields} for row in reader]


def write_csv(path: Path, fields: list[str], rows: Iterable[dict]) -> None:
    ensure_dirs()
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def append_csv(path: Path, fields: list[str], rows: list[dict]) -> None:
    existing = read_csv(path, fields)
    write_csv(path, fields, [*existing, *rows])


def upsert_daily_status(status_row: dict) -> None:
    existing = read_csv(DAILY_STATUS_CSV, DAILY_STATUS_FIELDS)
    remaining = [
        row
        for row in existing
        if row.get("snapshot_date") != status_row.get("snapshot_date")
    ]
    write_csv(DAILY_STATUS_CSV, DAILY_STATUS_FIELDS, [*remaining, status_row])


def expected_sheet_names() -> set[str]:
    return {"说明", "全部明细", *{f"{source.platform}-{source.chart}" for source in CHARTS}}


def workbook_has_expected_sheets(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as workbook:
            root = ET.fromstring(workbook.read("xl/workbook.xml"))
        ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets_node = root.find("a:sheets", ns)
        if sheets_node is None:
            return False
        sheet_names = {
            sheet.attrib.get("name", "")
            for sheet in sheets_node
        }
        return expected_sheet_names().issubset(sheet_names)
    except Exception:
        return False


def is_day_complete(snapshot_date: str) -> bool:
    expected = OUTPUT_DIR / f"music_charts_{snapshot_date}.xlsx"
    for row in read_csv(DAILY_STATUS_CSV, DAILY_STATUS_FIELDS):
        if (
            row.get("snapshot_date") == snapshot_date
            and row.get("status") == "complete"
            and expected.exists()
            and workbook_has_expected_sheets(expected)
        ):
            return True
    return False


def completed_dates(start_on: str = "") -> set[str]:
    dates = set()
    for row in read_csv(DAILY_STATUS_CSV, DAILY_STATUS_FIELDS):
        snapshot_date = row.get("snapshot_date", "")
        expected = OUTPUT_DIR / f"music_charts_{snapshot_date}.xlsx"
        if (
            snapshot_date
            and row.get("status") == "complete"
            and (not start_on or snapshot_date >= start_on)
            and expected.exists()
            and workbook_has_expected_sheets(expected)
        ):
            dates.add(snapshot_date)
    return dates


def configured_target_days(config: dict) -> int:
    try:
        return int(config.get("target_days", TARGET_DAYS))
    except (TypeError, ValueError):
        return TARGET_DAYS


def has_day_limit(config: dict) -> bool:
    return configured_target_days(config) > 0


def target_days_text(config: dict) -> str:
    return str(configured_target_days(config)) if has_day_limit(config) else "不限制"


def qq_api_url(top_id: str) -> str:
    payload = {
        "detail": {
            "module": "musicToplist.ToplistInfoServer",
            "method": "GetDetail",
            "param": {
                "topId": int(top_id),
                "offset": 0,
                "num": 100,
                "period": "",
            },
        }
    }
    encoded = urllib.parse.quote(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    )
    return f"https://u.y.qq.com/cgi-bin/musicu.fcg?data={encoded}"


def qq_metric(rank_type: int | None, rank_value: str) -> str:
    if rank_type == 1:
        return f"上升 {rank_value}".strip()
    if rank_type == 2:
        return f"下降 {rank_value}".strip()
    if rank_type == 4:
        return "新进"
    if rank_type == 6:
        return f"热度 {rank_value}".strip()
    return rank_value


def collect_qq(source: ChartSource, snapshot_date: str, captured_at: str) -> list[dict]:
    payload = json.loads(fetch_text(qq_api_url(source.rank_id)))
    detail = payload.get("detail", {}).get("data", {}).get("data", {})
    songs = detail.get("song") or []
    platform_date = str(detail.get("updateTime") or detail.get("period") or "")

    if len(songs) < source.expected_rows:
        raise RuntimeError(
            f"{source.platform}{source.chart}只返回 {len(songs)} 条，少于预期 {source.expected_rows} 条"
        )

    rows = []
    for item in songs[: source.expected_rows]:
        rows.append(
            {
                "snapshot_date": snapshot_date,
                "captured_at": captured_at,
                "platform": source.platform,
                "chart": source.chart,
                "official_chart": source.official_chart,
                "platform_date": platform_date,
                "rank": str(item.get("rank", "")),
                "song_name": html.unescape(str(item.get("title", ""))).strip(),
                "artist": html.unescape(str(item.get("singerName", ""))).strip(),
                "chart_metric": qq_metric(
                    item.get("rankType"), str(item.get("rankValue", ""))
                ),
                "duration": "",
                "song_id": str(item.get("songId", "")),
                "album_id": str(item.get("albumMid", "")),
                "source_url": source.url,
            }
        )
    return rows


def kugou_page_url(rank_id: str, page_num: int) -> str:
    return f"https://www.kugou.com/yy/rank/home/{page_num}-{rank_id}.html?from=rank"


def extract_kugou_platform_date(page: str) -> str:
    match = re.search(r"榜单更新于：([0-9-]+)", page)
    return match.group(1) if match else ""


def extract_kugou_features(page: str) -> list[dict]:
    match = re.search(r"global\.features\s*=\s*(\[.*?\]);", page, re.S)
    if not match:
        raise RuntimeError("未找到酷狗 global.features 数据")
    return json.loads(match.group(1))


def extract_kugou_page_info(page: str) -> tuple[int, int]:
    total_match = re.search(r"total:\s*'([0-9]+)'", page)
    page_size_match = re.search(r"pagesize:\s*'([0-9]+)'", page)
    if not total_match or not page_size_match:
        raise RuntimeError("未找到酷狗分页信息")
    return int(total_match.group(1)), int(page_size_match.group(1))


def split_kugou_name(file_name: str, fallback_artist: str) -> tuple[str, str]:
    decoded = html.unescape(file_name).strip()
    if " - " not in decoded:
        return html.unescape(fallback_artist).strip(), decoded
    artist, title = decoded.split(" - ", 1)
    return artist.strip(), title.strip()


def seconds_to_duration(value: object) -> str:
    try:
        seconds = int(round(float(value)))
    except (TypeError, ValueError):
        return ""
    return f"{seconds // 60}:{seconds % 60:02d}"


def collect_kugou(source: ChartSource, snapshot_date: str, captured_at: str) -> list[dict]:
    first_page = fetch_text(source.url)
    platform_date = extract_kugou_platform_date(first_page)
    total, page_size = extract_kugou_page_info(first_page)
    page_count = max(1, math.ceil(total / page_size))
    rows: list[dict] = []

    for page_num in range(1, page_count + 1):
        page_url = kugou_page_url(source.rank_id, page_num)
        page = first_page if page_num == 1 else fetch_text(page_url)
        for index, item in enumerate(extract_kugou_features(page), start=1):
            artist, song_name = split_kugou_name(
                str(item.get("FileName", "")),
                str(item.get("author_name", "")),
            )
            rows.append(
                {
                    "snapshot_date": snapshot_date,
                    "captured_at": captured_at,
                    "platform": source.platform,
                    "chart": source.chart,
                    "official_chart": source.official_chart,
                    "platform_date": platform_date,
                    "rank": str((page_num - 1) * page_size + index),
                    "song_name": song_name,
                    "artist": artist,
                    "chart_metric": "",
                    "duration": seconds_to_duration(item.get("timeLen")),
                    "song_id": str(item.get("Hash", "")),
                    "album_id": str(item.get("album_id", "")),
                    "source_url": page_url,
                }
            )
        time.sleep(0.2)

    rows = rows[:total]
    if len(rows) < source.expected_rows:
        raise RuntimeError(
            f"{source.platform}{source.chart}只返回 {len(rows)} 条，少于预期 {source.expected_rows} 条"
        )
    return rows[: source.expected_rows]


ENTERTAINMENT_CATEGORIES = {
    "艺人",
    "剧集",
    "电影",
    "综艺",
    "音乐",
    "演出",
    "盛典",
    "动漫",
}


def weibo_search_url(word: str) -> str:
    query = f"#{word}#"
    return "https://s.weibo.com/weibo?q=" + urllib.parse.quote(query)


def fetch_weibo_band_items(source: ChartSource) -> list[dict]:
    payload = json.loads(fetch_text(source.url))
    if payload.get("ok") != 1:
        raise RuntimeError(f"微博接口返回异常：{payload.get('ok')}")
    items = payload.get("data", {}).get("band_list") or []
    clean_items = []
    for item in items:
        if item.get("is_ad"):
            continue
        if not item.get("realpos"):
            continue
        word = str(item.get("word") or item.get("note") or "").strip()
        if not word:
            continue
        clean_items.append(item)
    clean_items.sort(key=lambda item: int(item.get("realpos") or 9999))
    return clean_items


def is_weibo_entertainment_item(item: dict) -> bool:
    category = str(item.get("category") or "").strip()
    channel_type = str(item.get("channel_type") or "").strip()
    return channel_type == "Entertainment" or category in ENTERTAINMENT_CATEGORIES


def weibo_metric(item: dict, total_rank: str | None = None) -> str:
    parts = []
    if total_rank:
        parts.append(f"总榜排名 {total_rank}")
    if item.get("num") not in (None, ""):
        parts.append(f"热度 {item.get('num')}")
    label = str(item.get("icon_desc") or item.get("label_name") or "").strip()
    if label:
        parts.append(label)
    field_tag = str(item.get("field_tag") or "").strip()
    if field_tag:
        parts.append(field_tag)
    return "；".join(parts)


def collect_weibo(source: ChartSource, snapshot_date: str, captured_at: str) -> list[dict]:
    items = fetch_weibo_band_items(source)
    if source.kind == "weibo_hot":
        selected = items[: source.expected_rows]
    else:
        selected = [item for item in items if is_weibo_entertainment_item(item)]
        selected = selected[: source.expected_rows]

    minimum_rows = source.minimum_rows or source.expected_rows
    if len(selected) < minimum_rows:
        raise RuntimeError(
            f"{source.platform}{source.chart}只返回 {len(selected)} 条，少于最低预期 {minimum_rows} 条"
        )

    rows = []
    for index, item in enumerate(selected, start=1):
        word = html.unescape(str(item.get("word") or item.get("note") or "")).strip()
        category = html.unescape(str(item.get("category") or "")).strip()
        total_rank = str(item.get("realpos") or "")
        rank = total_rank if source.kind == "weibo_hot" else str(index)
        rows.append(
            {
                "snapshot_date": snapshot_date,
                "captured_at": captured_at,
                "platform": source.platform,
                "chart": source.chart,
                "official_chart": source.official_chart,
                "platform_date": "",
                "rank": rank,
                "song_name": word,
                "artist": category,
                "chart_metric": weibo_metric(
                    item,
                    total_rank if source.kind == "weibo_entertainment" else None,
                ),
                "duration": "",
                "song_id": word,
                "album_id": str(item.get("subject_label") or item.get("star_name") or ""),
                "source_url": weibo_search_url(word),
            }
        )
    return rows


def collect_chart(source: ChartSource, snapshot_date: str, captured_at: str) -> list[dict]:
    if source.kind == "qq":
        return collect_qq(source, snapshot_date, captured_at)
    if source.kind == "kugou":
        return collect_kugou(source, snapshot_date, captured_at)
    if source.kind in {"weibo_hot", "weibo_entertainment"}:
        return collect_weibo(source, snapshot_date, captured_at)
    raise RuntimeError(f"未知榜单类型：{source.kind}")


def official_snapshot_date(rows: list[dict], fallback_date: str) -> str:
    dates_by_chart: dict[tuple[str, str], set[str]] = defaultdict(set)
    for row in rows:
        platform_date = str(row.get("platform_date", "")).strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", platform_date):
            dates_by_chart[(row["platform"], row["chart"])].add(platform_date)

    if not dates_by_chart:
        return fallback_date

    all_dates = sorted({date_value for values in dates_by_chart.values() for date_value in values})
    if len(all_dates) == 1:
        return all_dates[0]

    parts = []
    for (platform, chart), date_values in sorted(dates_by_chart.items()):
        parts.append(f"{platform}-{chart}: {', '.join(sorted(date_values))}")
    raise RuntimeError("每日更新榜单的平台显示日期不一致，等待后续候补触发重试：" + "；".join(parts))


def replace_snapshot_date(rows: list[dict], snapshot_date: str) -> None:
    for row in rows:
        row["snapshot_date"] = snapshot_date


def minimum_expected_rows(source: ChartSource) -> int:
    return source.minimum_rows or source.expected_rows


def col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xml_text(value: object) -> str:
    return escape(str(value), {'"': "&quot;"})


def cell_style(row_index: int) -> int:
    if row_index == 1:
        return 1
    if row_index == 2:
        return 2
    if row_index == 4:
        return 3
    return 0


def sheet_xml(
    rows: list[list[object]],
    widths: list[int],
    freeze_rows: int = 0,
    autofilter_row: int | None = None,
) -> str:
    max_col = max((len(row) for row in rows), default=1)
    max_row = len(rows)
    dimension = f"A1:{col_name(max_col - 1)}{max_row}"

    col_parts = ["<cols>"]
    for i, width in enumerate(widths, start=1):
        col_parts.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    col_parts.append("</cols>")

    views = '<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
    if freeze_rows:
        views += (
            f'<pane ySplit="{freeze_rows}" topLeftCell="A{freeze_rows + 1}" '
            'activePane="bottomLeft" state="frozen"/>'
        )
    views += "</sheetView></sheetViews>"

    sheet_data = ["<sheetData>"]
    for row_index, row in enumerate(rows, start=1):
        sheet_data.append(f'<row r="{row_index}">')
        for col_index, value in enumerate(row):
            if value is None or value == "":
                continue
            cell_ref = f"{col_name(col_index)}{row_index}"
            style_id = cell_style(row_index)
            style_attr = f' s="{style_id}"' if style_id else ""
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                sheet_data.append(f'<c r="{cell_ref}"{style_attr}><v>{value}</v></c>')
            else:
                sheet_data.append(
                    f'<c r="{cell_ref}" t="inlineStr"{style_attr}><is><t>{xml_text(value)}</t></is></c>'
                )
        sheet_data.append("</row>")
    sheet_data.append("</sheetData>")

    filter_xml = ""
    if autofilter_row and max_row >= autofilter_row:
        filter_xml = f'<autoFilter ref="A{autofilter_row}:{col_name(max_col - 1)}{max_row}"/>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        f"{views}"
        f"{''.join(col_parts)}"
        f"{''.join(sheet_data)}"
        f"{filter_xml}"
        "</worksheet>"
    )


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for index, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{xml_text(name)}" sheetId="{index}" r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheets)}</sheets>"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(rels)}"
        "</Relationships>"
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def content_types_xml(sheet_count: int) -> str:
    sheet_parts = []
    for index in range(1, sheet_count + 1):
        sheet_parts.append(
            '<Override '
            f'PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{''.join(sheet_parts)}"
        "</Types>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="4">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '<font><sz val="10"/><color rgb="FF17324D"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF17324D"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2B5C8A"/></patternFill></fill>'
        "</fills>"
        '<borders count="2">'
        "<border><left/><right/><top/><bottom/><diagonal/></border>"
        '<border><left style="thin"><color rgb="FFD7DEE8"/></left>'
        '<right style="thin"><color rgb="FFD7DEE8"/></right>'
        '<top style="thin"><color rgb="FFD7DEE8"/></top>'
        '<bottom style="thin"><color rgb="FFD7DEE8"/></bottom><diagonal/></border>'
        "</borders>"
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def write_xlsx(
    output_path: Path,
    sheet_map: dict[str, tuple[list[list[object]], list[int], int, int | None]],
) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    names = list(sheet_map.keys())
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml(len(names)))
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml(names))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(names)))
        archive.writestr("xl/styles.xml", styles_xml())
        for index, name in enumerate(names, start=1):
            rows, widths, freeze_rows, autofilter_row = sheet_map[name]
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                sheet_xml(rows, widths, freeze_rows, autofilter_row),
            )


def with_title(title: str, subtitle: str, headers: list[str], body: list[list[object]]) -> list[list[object]]:
    return [[title], [subtitle], [], headers, *body]


def detail_table(rows: list[dict]) -> list[list[object]]:
    headers = [
        "快照日期",
        "抓取时间",
        "平台",
        "榜单",
        "实际来源榜单",
        "平台显示日期",
        "排名",
        "条目/歌曲名",
        "歌手/分类",
        "指标/热度",
        "时长",
        "歌曲/Hash/词条ID",
        "专辑ID/MID/扩展",
        "来源链接",
    ]
    body = []
    for row in rows:
        body.append(
            [
                row["snapshot_date"],
                row["captured_at"],
                row["platform"],
                row["chart"],
                row["official_chart"],
                row["platform_date"],
                int(row["rank"]) if str(row["rank"]).isdigit() else row["rank"],
                row["song_name"],
                row["artist"],
                row["chart_metric"],
                row["duration"],
                row["song_id"],
                row["album_id"],
                row["source_url"],
            ]
        )
    return with_title(
        "榜单明细",
        "本表是一日快照：一行代表一个榜单条目在该日对应平台榜单里的排名。",
        headers,
        body,
    )


def rows_for_chart(rows: list[dict], source: ChartSource) -> list[dict]:
    return [
        row
        for row in rows
        if row["platform"] == source.platform and row["chart"] == source.chart
    ]


def build_daily_workbook(snapshot_date: str, captured_at: str, rows: list[dict]) -> Path:
    output_path = OUTPUT_DIR / f"music_charts_{snapshot_date}.xlsx"
    summary_headers = ["项目", "内容"]
    summary_body = [
        ["快照日期", snapshot_date],
        ["抓取时间", captured_at],
        ["总行数", len(rows)],
        ["输出文件", str(output_path.as_posix())],
        ["说明", "每天多次候补触发，但同一日期只生成一张完整xlsx表格。微博热搜为实时榜，按采集时间记录。"],
    ]
    for source in CHARTS:
        count = len(rows_for_chart(rows, source))
        summary_body.append(
            [
                f"{source.platform} {source.chart}",
                f"{count} 行；来源口径：{source.official_chart}；{source.note}",
            ]
        )

    sheet_map: dict[str, tuple[list[list[object]], list[int], int, int | None]] = {
        "说明": (
            with_title(
                "每日榜单快照",
                "QQ音乐、酷狗音乐、微博热搜，每天一张表。",
                summary_headers,
                summary_body,
            ),
            [24, 110],
            4,
            4,
        ),
        "全部明细": (
            detail_table(rows),
            [13, 24, 12, 12, 16, 16, 8, 28, 28, 14, 10, 34, 18, 58],
            4,
            4,
        ),
    }

    for source in CHARTS:
        chart_rows = rows_for_chart(rows, source)
        sheet_map[f"{source.platform}-{source.chart}"] = (
            detail_table(chart_rows),
            [13, 24, 12, 12, 16, 16, 8, 28, 28, 14, 10, 34, 18, 58],
            4,
            4,
        )

    write_xlsx(output_path, sheet_map)
    return output_path


def write_run_status(config: dict, message: str) -> None:
    statuses = read_csv(DAILY_STATUS_CSV, DAILY_STATUS_FIELDS)
    collection_start_on = str(config.get("collection_start_on", "")).strip()
    complete = []
    for row in statuses:
        snapshot_date = row.get("snapshot_date", "")
        expected = OUTPUT_DIR / f"music_charts_{snapshot_date}.xlsx"
        if (
            snapshot_date
            and row.get("status") == "complete"
            and (not collection_start_on or snapshot_date >= collection_start_on)
            and expected.exists()
            and workbook_has_expected_sheets(expected)
        ):
            complete.append(row)
    lines = [
        "# 音乐榜单每日采集状态",
        "",
        f"- 状态：{message}",
        f"- 采集起始日期：{collection_start_on}",
        f"- 目标天数：{target_days_text(config)}",
        f"- 已完成天数：{len(complete)}",
        f"- 首次运行日期：{config.get('started_on', '')}",
        "",
        "## 已生成表格",
        "",
    ]
    if complete:
        for row in sorted(complete, key=lambda item: item["snapshot_date"]):
            lines.append(
                f"- {row['snapshot_date']}: [{Path(row['xlsx_path']).name}]({row['xlsx_path']})"
            )
    else:
        lines.append("- 暂无")
    RUN_STATUS_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    ensure_dirs()
    started = now_local()
    today = started.date()
    snapshot_date = today.isoformat()
    captured_at = iso(started)
    run_id = started.strftime("%Y%m%d_%H%M%S")
    config = load_or_create_config(today)
    collection_start_on = str(config.get("collection_start_on", "")).strip()

    if has_day_limit(config) and len(completed_dates(collection_start_on)) >= configured_target_days(config):
        write_run_status(config, f"已完成{configured_target_days(config)}天采集，本次自动跳过。")
        print("已完成目标天数，本次跳过。")
        return 0

    all_rows: list[dict] = []
    heartbeats: list[dict] = []

    for source in CHARTS:
        chart_started = now_local()
        status = "success"
        rows: list[dict] = []
        error = ""
        try:
            rows = collect_chart(source, snapshot_date, captured_at)
            all_rows.extend(rows)
        except Exception as exc:
            status = "failed"
            error = str(exc)[:500]
        finally:
            chart_finished = now_local()
            heartbeats.append(
                {
                    "run_id": run_id,
                    "snapshot_date": snapshot_date,
                    "status": status,
                    "platform": source.platform,
                    "chart": source.chart,
                    "official_chart": source.official_chart,
                    "row_count": str(len(rows)),
                    "started_at": iso(chart_started),
                    "finished_at": iso(chart_finished),
                    "error_message": error,
                }
            )

    success_count = sum(1 for row in heartbeats if row["status"] == "success")

    if success_count != len(CHARTS):
        append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
        message = f"{snapshot_date} 本次只成功 {success_count}/{len(CHARTS)} 个榜单，等待后续候补触发重试。"
        upsert_daily_status(
            {
                "snapshot_date": snapshot_date,
                "status": "failed",
                "row_count": str(len(all_rows)),
                "xlsx_path": "",
                "completed_at": "",
                "message": message,
            }
        )
        write_run_status(config, message)
        print(message)
        return 1

    expected_min_total = sum(minimum_expected_rows(source) for source in CHARTS)
    if len(all_rows) < expected_min_total:
        append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
        message = f"{snapshot_date} 行数异常：实际 {len(all_rows)}，最低预期 {expected_min_total}。"
        upsert_daily_status(
            {
                "snapshot_date": snapshot_date,
                "status": "failed",
                "row_count": str(len(all_rows)),
                "xlsx_path": "",
                "completed_at": "",
                "message": message,
            }
        )
        write_run_status(config, message)
        print(message)
        return 1

    try:
        snapshot_date = official_snapshot_date(all_rows, snapshot_date)
    except Exception as exc:
        append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
        message = str(exc)[:500]
        upsert_daily_status(
            {
                "snapshot_date": today.isoformat(),
                "status": "failed",
                "row_count": str(len(all_rows)),
                "xlsx_path": "",
                "completed_at": "",
                "message": message,
            }
        )
        write_run_status(config, message)
        print(message)
        return 1

    if collection_start_on and snapshot_date < collection_start_on:
        replace_snapshot_date(heartbeats, snapshot_date)
        append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
        message = f"{snapshot_date} 早于采集起始日期 {collection_start_on}，本次自动跳过。"
        write_run_status(config, message)
        print(message)
        return 0

    replace_snapshot_date(all_rows, snapshot_date)
    replace_snapshot_date(heartbeats, snapshot_date)

    if is_day_complete(snapshot_date):
        append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
        write_run_status(config, f"{snapshot_date} 已经有完整表格，本次候补触发自动跳过。")
        print(f"{snapshot_date} 已完成，本次跳过。")
        return 0

    append_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, heartbeats)
    detail_csv = DATA_DIR / f"daily_rows_{snapshot_date}.csv"
    write_csv(detail_csv, DETAIL_FIELDS, all_rows)
    output_path = build_daily_workbook(snapshot_date, captured_at, all_rows)
    message = f"{snapshot_date} 完成，生成 {output_path.name}，共 {len(all_rows)} 行。"
    upsert_daily_status(
        {
            "snapshot_date": snapshot_date,
            "status": "complete",
            "row_count": str(len(all_rows)),
            "xlsx_path": output_path.as_posix(),
            "completed_at": captured_at,
            "message": message,
        }
    )
    write_run_status(config, message)
    print(message)
    return 0


if __name__ == "__main__":
    sys.exit(main())
