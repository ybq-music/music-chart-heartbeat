from __future__ import annotations

import csv
import html
import json
import math
import os
import re
import sys
import time
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "outputs"
CONFIG_PATH = DATA_DIR / "run_config.json"
ROWS_CSV = DATA_DIR / "chart_rows.csv"
HEARTBEAT_CSV = DATA_DIR / "chart_heartbeat.csv"
WORKBOOK_PATH = OUTPUT_DIR / "music_chart_heartbeat.xlsx"
TZ = ZoneInfo("Asia/Shanghai")

RUN_DAYS = int(os.getenv("RUN_DAYS", "7"))
REQUEST_TIMEOUT_SECONDS = int(os.getenv("REQUEST_TIMEOUT_SECONDS", "30"))


ROW_FIELDS = [
    "snapshot_id",
    "scheduled_at",
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
    "scheduled_at",
    "platform",
    "chart",
    "official_chart",
    "status",
    "row_count",
    "started_at",
    "finished_at",
    "error_message",
]


@dataclass(frozen=True)
class ChartSource:
    platform: str
    chart: str
    official_chart: str
    kind: str
    url: str
    rank_id: str = ""
    note: str = ""


CHARTS = [
    ChartSource(
        platform="QQ音乐",
        chart="热歌榜",
        official_chart="热歌榜",
        kind="qq",
        url="https://y.qq.com/n/ryqq/toplist/26",
        note="QQ音乐公开网页榜单当前展示20条。",
    ),
    ChartSource(
        platform="QQ音乐",
        chart="飙升榜",
        official_chart="飙升榜",
        kind="qq",
        url="https://y.qq.com/n/ryqq/toplist/62",
        note="QQ音乐公开网页榜单当前展示20条。",
    ),
    ChartSource(
        platform="酷狗音乐",
        chart="热歌榜",
        official_chart="酷狗TOP500",
        kind="kugou",
        rank_id="8888",
        url="https://www.kugou.com/yy/rank/home/1-8888.html?from=rank",
        note="酷狗公开榜单没有直接命名为“热歌榜”的入口，本任务采用官方“酷狗TOP500”作为热歌榜口径。",
    ),
    ChartSource(
        platform="酷狗音乐",
        chart="飙升榜",
        official_chart="酷狗飙升榜",
        kind="kugou",
        rank_id="6666",
        url="https://www.kugou.com/yy/rank/home/1-6666.html?from=rank",
        note="酷狗飙升榜按分页抓取。",
    ),
]


def now_local() -> datetime:
    return datetime.now(TZ)


def iso(dt: datetime) -> str:
    return dt.replace(microsecond=0).isoformat()


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def load_or_create_config(started_at: datetime) -> dict:
    ensure_dirs()
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

    ends_at = started_at + timedelta(days=RUN_DAYS)
    config = {
        "timezone": "Asia/Shanghai",
        "run_days": RUN_DAYS,
        "started_at": iso(started_at),
        "ends_at": iso(ends_at),
    }
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return config


def fetch_text(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0 Safari/537.36"
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


def extract_platform_date(page: str) -> str:
    qq_match = re.search(r'toplist_switch__data">([^<]+)<', page)
    if qq_match:
        return html.unescape(qq_match.group(1)).strip()

    kugou_match = re.search(r"榜单更新于：([0-9-]+)", page)
    if kugou_match:
        return kugou_match.group(1)

    return ""


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


def collect_qq(source: ChartSource, scheduled_at: str, captured_at: str) -> list[dict]:
    page = fetch_text(source.url)
    platform_date = extract_platform_date(page)
    match = re.search(r'"rankList":(\[.*?\]),"historyarr"', page, re.S)
    if not match:
        raise RuntimeError("未找到 QQ 音乐 rankList 数据")

    rank_list = json.loads(match.group(1))
    rows = []
    snapshot_id = snapshot_key(scheduled_at, source)
    for item in rank_list:
        rows.append(
            {
                "snapshot_id": snapshot_id,
                "scheduled_at": scheduled_at,
                "captured_at": captured_at,
                "platform": source.platform,
                "chart": source.chart,
                "official_chart": source.official_chart,
                "platform_date": platform_date,
                "rank": str(item.get("rank", "")),
                "song_name": html.unescape(str(item.get("title", ""))).strip(),
                "artist": html.unescape(str(item.get("singerName", ""))).strip(),
                "chart_metric": qq_metric(item.get("rankType"), str(item.get("rankValue", ""))),
                "duration": "",
                "song_id": str(item.get("songId", "")),
                "album_id": str(item.get("albumMid", "")),
                "source_url": source.url,
            }
        )
    return rows


def kugou_page_url(rank_id: str, page_num: int) -> str:
    return f"https://www.kugou.com/yy/rank/home/{page_num}-{rank_id}.html?from=rank"


def extract_kugou_features(page: str) -> list[dict]:
    match = re.search(r"global\.features\s*=\s*(\[.*?\]);", page, re.S)
    if not match:
        raise RuntimeError("未找到酷狗 global.features 数据")
    return json.loads(match.group(1))


def extract_kugou_page_info(page: str) -> tuple[int, int]:
    total = int(re.search(r"total:\s*'([0-9]+)'", page).group(1))
    page_size = int(re.search(r"pagesize:\s*'([0-9]+)'", page).group(1))
    return total, page_size


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


def collect_kugou(source: ChartSource, scheduled_at: str, captured_at: str) -> list[dict]:
    first_page = fetch_text(source.url)
    platform_date = extract_platform_date(first_page)
    total, page_size = extract_kugou_page_info(first_page)
    page_count = max(1, math.ceil(total / page_size))
    rows: list[dict] = []
    snapshot_id = snapshot_key(scheduled_at, source)

    for page_num in range(1, page_count + 1):
        page_url = kugou_page_url(source.rank_id, page_num)
        page = first_page if page_num == 1 else fetch_text(page_url)
        features = extract_kugou_features(page)

        for index, item in enumerate(features, start=1):
            artist, song_name = split_kugou_name(
                str(item.get("FileName", "")),
                str(item.get("author_name", "")),
            )
            rows.append(
                {
                    "snapshot_id": snapshot_id,
                    "scheduled_at": scheduled_at,
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

    return rows[:total]


def snapshot_key(scheduled_at: str, source: ChartSource) -> str:
    compact = scheduled_at.replace("-", "").replace(":", "").replace("+08:00", "")
    compact = compact.replace("T", "_")
    return f"{compact}_{source.platform}_{source.chart}"


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


def append_dedup_rows(new_rows: list[dict]) -> None:
    existing = read_csv(ROWS_CSV, ROW_FIELDS)
    seen = {
        (
            row["scheduled_at"],
            row["platform"],
            row["chart"],
            row["rank"],
            row["song_id"],
        )
        for row in existing
    }

    merged = list(existing)
    for row in new_rows:
        key = (
            row["scheduled_at"],
            row["platform"],
            row["chart"],
            row["rank"],
            row["song_id"],
        )
        if key not in seen:
            merged.append(row)
            seen.add(key)

    write_csv(ROWS_CSV, ROW_FIELDS, merged)


def append_heartbeats(new_rows: list[dict]) -> None:
    existing = read_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS)
    write_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS, [*existing, *new_rows])


def col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xml_text(value: object) -> str:
    return escape(str(value), {'"': "&quot;"})


def sheet_xml(
    rows: list[list[object]],
    widths: list[int] | None = None,
    freeze_rows: int = 0,
    autofilter_row: int | None = None,
) -> str:
    widths = widths or []
    max_col = max((len(row) for row in rows), default=1)
    max_row = len(rows)
    dimension = f"A1:{col_name(max_col - 1)}{max_row}"

    col_xml = ""
    if widths:
        parts = ["<cols>"]
        for i, width in enumerate(widths, start=1):
            parts.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
        parts.append("</cols>")
        col_xml = "".join(parts)

    views = '<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
    if freeze_rows:
        views += (
            f'<pane ySplit="{freeze_rows}" topLeftCell="A{freeze_rows + 1}" '
            'activePane="bottomLeft" state="frozen"/>'
        )
    views += "</sheetView></sheetViews>"

    row_xml_parts = ["<sheetData>"]
    for row_index, row in enumerate(rows, start=1):
        row_xml_parts.append(f'<row r="{row_index}">')
        for col_index, value in enumerate(row):
            if value is None or value == "":
                continue
            cell_ref = f"{col_name(col_index)}{row_index}"
            style_id = cell_style(row_index)
            style_attr = f' s="{style_id}"' if style_id else ""
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                row_xml_parts.append(
                    f'<c r="{cell_ref}"{style_attr}><v>{value}</v></c>'
                )
            else:
                row_xml_parts.append(
                    f'<c r="{cell_ref}" t="inlineStr"{style_attr}><is><t>{xml_text(value)}</t></is></c>'
                )
        row_xml_parts.append("</row>")
    row_xml_parts.append("</sheetData>")

    filter_xml = ""
    if autofilter_row and max_row >= autofilter_row:
        filter_xml = f'<autoFilter ref="A{autofilter_row}:{col_name(max_col - 1)}{max_row}"/>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        f"{views}"
        f"{col_xml}"
        f"{''.join(row_xml_parts)}"
        f"{filter_xml}"
        "</worksheet>"
    )


def cell_style(row_index: int) -> int:
    if row_index == 1:
        return 1
    if row_index == 2:
        return 2
    if row_index == 4:
        return 3
    return 0


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
    sheets = []
    for index in range(1, sheet_count + 1):
        sheets.append(
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
        f"{''.join(sheets)}"
        "</Types>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<fonts count=\"4\">"
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '<font><sz val="10"/><color rgb="FF17324D"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        "</fonts>"
        "<fills count=\"4\">"
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF17324D"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2B5C8A"/></patternFill></fill>'
        "</fills>"
        "<borders count=\"2\">"
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


def write_xlsx(sheet_map: dict[str, tuple[list[list[object]], list[int], int, int | None]]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    names = list(sheet_map.keys())
    with zipfile.ZipFile(WORKBOOK_PATH, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml(len(names)))
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml(names))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(names)))
        archive.writestr("xl/styles.xml", styles_xml())
        for index, name in enumerate(names, start=1):
            rows, widths, freeze_rows, autofilter_row = sheet_map[name]
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                sheet_xml(rows, widths=widths, freeze_rows=freeze_rows, autofilter_row=autofilter_row),
            )


def chart_display_rows(rows: list[dict]) -> list[list[object]]:
    headers = [
        "抓取时间",
        "平台",
        "榜单",
        "实际来源榜单",
        "排名",
        "歌曲名",
        "歌手",
        "榜单指标",
        "时长",
        "歌曲/Hash ID",
        "专辑ID/MID",
        "平台显示日期",
        "来源链接",
    ]
    body = [
        [
            row["captured_at"],
            row["platform"],
            row["chart"],
            row["official_chart"],
            int(row["rank"]) if str(row["rank"]).isdigit() else row["rank"],
            row["song_name"],
            row["artist"],
            row["chart_metric"],
            row["duration"],
            row["song_id"],
            row["album_id"],
            row["platform_date"],
            row["source_url"],
        ]
        for row in rows
    ]
    return body_with_title("榜单明细", "每一行是一首歌在某一次每小时快照中的排名。", headers, body)


def body_with_title(
    title: str,
    subtitle: str,
    headers: list[str],
    body: list[list[object]],
) -> list[list[object]]:
    return [[title], [subtitle], [], headers, *body]


def latest_rows_by_chart(all_rows: list[dict], source: ChartSource) -> list[dict]:
    matched = [
        row
        for row in all_rows
        if row["platform"] == source.platform and row["chart"] == source.chart
    ]
    if not matched:
        return []
    latest = max(row["scheduled_at"] for row in matched)
    return [row for row in matched if row["scheduled_at"] == latest]


def build_workbook(config: dict) -> None:
    all_rows = read_csv(ROWS_CSV, ROW_FIELDS)
    heartbeats = read_csv(HEARTBEAT_CSV, HEARTBEAT_FIELDS)

    summary_headers = ["项目", "内容"]
    latest_scheduled_at = max((row["scheduled_at"] for row in all_rows), default="")
    summary_body = [
        ["首次运行", config.get("started_at", "")],
        ["自动停止时间", config.get("ends_at", "")],
        ["最近快照小时", latest_scheduled_at],
        ["总明细行数", len(all_rows)],
        ["心跳记录数", len(heartbeats)],
        ["输出文件", str(WORKBOOK_PATH.as_posix())],
    ]
    for source in CHARTS:
        latest_count = len(latest_rows_by_chart(all_rows, source))
        summary_body.append(
            [
                f"{source.platform} {source.chart}",
                f"最近一次 {latest_count} 行；来源口径：{source.official_chart}；{source.note}",
            ]
        )

    heartbeat_headers = HEARTBEAT_FIELDS
    heartbeat_body = [[row[field] for field in HEARTBEAT_FIELDS] for row in heartbeats]

    sheet_map: dict[str, tuple[list[list[object]], list[int], int, int | None]] = {
        "说明": (
            body_with_title(
                "音乐榜单心跳任务",
                "自动抓取 QQ音乐、酷狗音乐热歌榜/飙升榜，并保留每小时历史快照。",
                summary_headers,
                summary_body,
            ),
            [24, 110],
            4,
            4,
        ),
        "运行心跳": (
            body_with_title(
                "运行心跳",
                "每个榜单每次运行一行，status=success 表示该榜单本小时抓取成功。",
                heartbeat_headers,
                heartbeat_body,
            ),
            [24, 24, 14, 14, 16, 12, 10, 24, 24, 80],
            4,
            4,
        ),
        "全部明细": (
            chart_display_rows(all_rows),
            [24, 12, 12, 16, 8, 28, 26, 14, 10, 34, 18, 16, 58],
            4,
            4,
        ),
    }

    for source in CHARTS:
        rows = latest_rows_by_chart(all_rows, source)
        sheet_name = f"{source.platform}-{source.chart}"
        sheet_map[sheet_name] = (
            chart_display_rows(rows),
            [24, 12, 12, 16, 8, 28, 26, 14, 10, 34, 18, 16, 58],
            4,
            4,
        )

    write_xlsx(sheet_map)


def collect_one(source: ChartSource, scheduled_at: str, captured_at: str) -> list[dict]:
    if source.kind == "qq":
        return collect_qq(source, scheduled_at, captured_at)
    if source.kind == "kugou":
        return collect_kugou(source, scheduled_at, captured_at)
    raise RuntimeError(f"未知榜单类型：{source.kind}")


def main() -> int:
    ensure_dirs()
    started = now_local()
    config = load_or_create_config(started)
    ends_at = datetime.fromisoformat(config["ends_at"])
    if started >= ends_at:
        print(f"任务已超过自动停止时间：{config['ends_at']}，本次不再抓取。")
        return 0

    scheduled = started.replace(minute=0, second=0, microsecond=0)
    scheduled_at = iso(scheduled)
    run_id = started.strftime("%Y%m%d_%H%M%S")
    captured_at = iso(started)

    all_new_rows: list[dict] = []
    heartbeats: list[dict] = []

    for source in CHARTS:
        chart_started = now_local()
        status = "success"
        rows: list[dict] = []
        error = ""
        try:
            rows = collect_one(source, scheduled_at, captured_at)
            all_new_rows.extend(rows)
        except Exception as exc:
            status = "failed"
            error = str(exc)[:500]
        finally:
            chart_finished = now_local()
            heartbeats.append(
                {
                    "run_id": run_id,
                    "scheduled_at": scheduled_at,
                    "platform": source.platform,
                    "chart": source.chart,
                    "official_chart": source.official_chart,
                    "status": status,
                    "row_count": str(len(rows)),
                    "started_at": iso(chart_started),
                    "finished_at": iso(chart_finished),
                    "error_message": error,
                }
            )

    append_dedup_rows(all_new_rows)
    append_heartbeats(heartbeats)
    build_workbook(config)

    success_count = sum(1 for row in heartbeats if row["status"] == "success")
    print(
        f"完成：{success_count}/{len(CHARTS)} 个榜单成功，"
        f"新增候选明细 {len(all_new_rows)} 行，输出 {WORKBOOK_PATH}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
