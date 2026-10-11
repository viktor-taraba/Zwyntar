"""
Збір даних з YouTube-каналу через YouTube Data API v3 та збереження в Excel.

Результат - один файл tsvyntar_youtube.xlsx з аркушами:
  - Пісні               : усі пісні - відео каналу та треки релізів, яких немає у завантаженнях
                          (колонка «Джерело»): назва, дата публікації, тривалість,
                          перегляди, лайки, к-ть коментарів, тематичні категорії,
                          звичайні плейлисти та релізи, до яких пісня входить
  - Плейлисти           : усі плейлисти каналу (колонка «Реліз?» позначає релізи)
  - Релізи              : альбоми/сингли із сумарними переглядами й лайками
  - Треки релізів       : треки кожного релізу зі статистикою
  - Коментарі по датах  : к-ть коментарів по датах для кожного відео
                          (верхній рівень, відповіді та сума)
  - Про дані            : канал, дата збору, примітки
"""

import os
import re
import sys
from collections import Counter
from datetime import date, datetime
from urllib.parse import unquote

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()
API_KEY = os.environ.get("YOUTUBE_API_KEY")
CHANNEL_ID = "UCio0I8HBJ7rWc7tk17rYqHw"
OUTPUT_FILE = "zwyntar_youtube.xlsx"

# True - у списку пісень залишити лише відео з музичною тематикою
ONLY_MUSIC = False
# True - рахувати коментарі по датах (1 одиниця квоти на кожні 100 коментарів)
COLLECT_COMMENTS = True
# True - рахувати також відповіді на коментарі (нижній рівень). Це додаткові запити:
# для гілок, де відповідей більше 5, викликається comments.list (1 одиниця за 100 відповідей)
INCLUDE_REPLIES = True

# Плейлисти-релізи YouTube Music мають ID, що починаються з "OLAK5uy_".
# Якщо API не віддасть їх серед плейлистів каналу, додайте ID вручну:
# відкрийте реліз на вкладці «Релізи» -> у посиланні буде ...?list=OLAK5uy_xxxxxxxx
RELEASE_PREFIX = "OLAK5uy_"
RELEASE_PLAYLIST_IDS = [
    "OLAK5uy_neyxxOb6OQlTQBRXedyouomcY1aScvhVQ",
    "OLAK5uy_lHAC-UUl7dq4fhBdU7S5RK10uG6wdyXJE",
    "OLAK5uy_lRNDRXqkM-ZoVuvrk8jAdfTxYnlur3YnY",
    "OLAK5uy_mOfBsfeuqouodl8nvgEmQ0_YEhfrbOQcg",
    "OLAK5uy_kw7_qIWs2jSEgzmDMnFfe7HtrPmR3wG_I",
    "OLAK5uy_kDCfQJ0dT3Y4MeHd-hN2EsCSJkOMLEt10",
    "OLAK5uy_kbQvULq_9zT8-u55527jBaXHb4JwVUSxA",
    "OLAK5uy_lp-HPRkkZkWh1BKKSQfcaHea0bbnoRmI4",
    "OLAK5uy_nRNes5qTyx2VQCXNcXRqi2nbK8e5VVpdw",
    "OLAK5uy_kdPYGpGTHPBNNwI6MGqfgZYWf51944uOg",
    "OLAK5uy_kTh8T8n4YDN5drotBb47BeetBZosg5BVc",
    "OLAK5uy_lmSJKkFGp9Wpl8TsK81TDxHCC91qMxJUM",
]

BASE_URL = "https://www.googleapis.com/youtube/v3"

FONT_NAME = "Arial"
DATE_FMT = "DD.MM.YYYY"
DATETIME_FMT = "DD.MM.YYYY HH:MM"
INT_FMT = "#,##0"
SOURCE_UPLOAD = "Завантаження каналу"
SOURCE_RELEASE_ONLY = "Лише в релізі"


class ApiError(Exception):
    def __init__(self, status, reason, message):
        super().__init__(f"HTTP {status} [{reason}]: {message}")
        self.status = status
        self.reason = reason


def api_get(endpoint, **params):
    """Один запит до API з обробкою помилок."""
    params = {**params, "key": API_KEY}
    resp = requests.get(f"{BASE_URL}/{endpoint}", params=params, timeout=30)
    if resp.status_code != 200:
        try:
            err = resp.json()["error"]
            reason = err["errors"][0].get("reason", "unknown")
            message = err.get("message", "")
        except Exception:
            reason, message = "unknown", resp.text[:200]
        raise ApiError(resp.status_code, reason, message)
    return resp.json()


def paginate(endpoint, **params):
    """Генератор, що проходить усі сторінки результатів."""
    token = None
    while True:
        if token:
            params["pageToken"] = token
        data = api_get(endpoint, **params)
        yield from data.get("items", [])
        token = data.get("nextPageToken")
        if not token:
            break


def chunks(seq, size):
    seq = list(seq)
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


# ------------------------------ Допоміжні -------------------------------
def parse_iso(value):
    """'2025-09-28T10:00:00Z' -> datetime (UTC, без часової зони, як вимагає Excel)."""
    return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)


def parse_duration(iso):
    """'PT3M45S' -> (225, '03:45')."""
    m = re.fullmatch(r"P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?", iso or "")
    if not m:
        return 0, ""
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    return seconds_to_text(d * 86400 + h * 3600 + mi * 60 + s)


def seconds_to_text(total):
    hh, rem = divmod(total, 3600)
    mm, ss = divmod(rem, 60)
    text = f"{hh:02d}:{mm:02d}:{ss:02d}" if hh else f"{mm:02d}:{ss:02d}"
    return total, text


def topic_name(url):
    """'https://en.wikipedia.org/wiki/Rock_music' -> 'Rock music'."""
    return unquote(url.rstrip("/").rsplit("/", 1)[-1]).replace("_", " ")


def to_int(value):
    """Статистика приходить рядками; поле може бути відсутнє (прихований лічильник)."""
    return int(value) if value is not None else ""


def is_release(playlist_id):
    return playlist_id.startswith(RELEASE_PREFIX) or playlist_id in RELEASE_PLAYLIST_IDS


def sum_numbers(values):
    return sum(v for v in values if v != "")


# --------------------------- Отримання даних ----------------------------
def get_channel():
    data = api_get("channels", part="snippet,contentDetails,statistics", id=CHANNEL_ID)
    items = data.get("items", [])
    if not items:
        sys.exit(f"Канал {CHANNEL_ID} не знайдено.")
    return items[0]


def get_all_video_ids(uploads_playlist_id):
    return [item["contentDetails"]["videoId"]
            for item in paginate("playlistItems", part="contentDetails",
                                 playlistId=uploads_playlist_id, maxResults=50)]


def get_video_details(video_ids):
    videos = []
    for batch in chunks(video_ids, 50):
        data = api_get("videos", part="snippet,contentDetails,statistics,topicDetails",
                       id=",".join(batch))
        videos.extend(data.get("items", []))
    return videos


def playlist_row(pl):
    return {
        "playlist_id": pl["id"],
        "title": pl["snippet"]["title"],
        "published_at": parse_iso(pl["snippet"]["publishedAt"]),
        "item_count": pl["contentDetails"]["itemCount"],
        "is_release": is_release(pl["id"]),
        "url": f"https://www.youtube.com/playlist?list={pl['id']}",
    }


def get_playlists():
    """Плейлисти каналу, які віддає API."""
    return [playlist_row(pl) for pl in paginate("playlists", part="snippet,contentDetails",
                                                channelId=CHANNEL_ID, maxResults=50)]


def get_playlists_by_id(playlist_ids):
    """Плейлисти за ID (для релізів, доданих вручну)."""
    rows = []
    for batch in chunks(playlist_ids, 50):
        try:
            data = api_get("playlists", part="snippet,contentDetails", id=",".join(batch))
            rows.extend(playlist_row(pl) for pl in data.get("items", []))
        except ApiError as e:
            print(f"  ! Не вдалося отримати плейлисти {batch}: {e}")
    return rows


def map_videos_to_playlists(playlists):
    """
    Повертає:
      mapping: video_id -> {"playlists": [...], "releases": [...]}
      items_by_playlist: playlist_id -> [video_id у порядку треків]
    """
    mapping, items_by_playlist = {}, {}
    for pl in playlists:
        kind = "releases" if pl["is_release"] else "playlists"
        ids = []
        try:
            for item in paginate("playlistItems", part="contentDetails",
                                 playlistId=pl["playlist_id"], maxResults=50):
                vid = item["contentDetails"]["videoId"]
                ids.append(vid)
                entry = mapping.setdefault(vid, {"playlists": [], "releases": []})
                entry[kind].append(pl["title"])
        except ApiError as e:
            print(f"  ! Плейлист '{pl['title']}' пропущено: {e}")
        items_by_playlist[pl["playlist_id"]] = ids
    return mapping, items_by_playlist


def fetch_all_replies(thread_id, fallback):
    """Усі відповіді в гілці (commentThreads віддає не більше 5 у відповіді)."""
    try:
        return list(paginate("comments", part="snippet", parentId=thread_id,
                             maxResults=100, textFormat="plainText"))
    except ApiError as e:
        print(f"    ! відповіді гілки {thread_id} отримано не повністю: {e}")
        return fallback


def count_comments_by_date(video_id):
    """
    Рахує коментарі по датах (UTC).
    Повертає (top_level: Counter, replies: Counter); replies порожній, якщо INCLUDE_REPLIES = False.
    """
    top, replies = Counter(), Counter()
    part = "snippet,replies" if INCLUDE_REPLIES else "snippet"
    try:
        for thread in paginate("commentThreads", part=part, videoId=video_id,
                               maxResults=100, order="time", textFormat="plainText"):
            snip = thread["snippet"]
            top[snip["topLevelComment"]["snippet"]["publishedAt"][:10]] += 1
            if not INCLUDE_REPLIES:
                continue
            total_replies = snip.get("totalReplyCount", 0)
            if total_replies == 0:
                continue
            embedded = thread.get("replies", {}).get("comments", [])
            items = embedded if len(embedded) >= total_replies \
                else fetch_all_replies(thread["id"], embedded)
            for rep in items:
                replies[rep["snippet"]["publishedAt"][:10]] += 1
    except ApiError as e:
        if e.reason == "commentsDisabled":
            print("    коментарі вимкнено")
        else:
            print(f"    ! помилка коментарів: {e}")
    return top, replies


# -------------------------------- Рядки ---------------------------------
def base_row(v):
    """Спільні поля для пісні/треку."""
    topics = [topic_name(t) for t in v.get("topicDetails", {}).get("topicCategories", [])]
    dur_sec, dur_text = parse_duration(v["contentDetails"].get("duration"))
    st = v.get("statistics", {})
    published = parse_iso(v["snippet"]["publishedAt"])
    return {
        "video_id": v["id"],
        "title": v["snippet"]["title"],
        "url": f"https://www.youtube.com/watch?v={v['id']}",
        "published_at": published,
        "published_date": published.date(),
        "duration_sec": dur_sec,
        "duration": dur_text,
        "views": to_int(st.get("viewCount")),
        "likes": to_int(st.get("likeCount")),
        "comments": to_int(st.get("commentCount")),
        "topic_categories": "; ".join(topics),
    }


def build_releases(playlists, items_by_playlist, details_by_id, upload_ids):
    """Повертає (releases_rows, release_tracks_rows)."""
    releases, tracks_rows = [], []
    for pl in playlists:
        if not pl["is_release"]:
            continue
        tracks, skipped = [], 0
        for no, vid in enumerate(items_by_playlist.get(pl["playlist_id"], []), 1):
            v = details_by_id.get(vid)
            if v is None:  # приватне/видалене відео
                skipped += 1
                continue
            tracks.append({
                "release_title": pl["title"],
                "track_no": no,
                **base_row(v),
                "in_channel_uploads": vid in upload_ids,
            })
        tracks_rows.extend(tracks)
        _, total_text = seconds_to_text(sum(t["duration_sec"] for t in tracks))
        releases.append({
            "release_id": pl["playlist_id"],
            "title": pl["title"],
            "published_at": pl["published_at"],
            "tracks": len(tracks),
            "tracks_unavailable": skipped,
            "type_guess": "сингл" if pl["item_count"] == 1 else "альбом/EP",
            "total_duration": total_text,
            "total_views": sum_numbers(t["views"] for t in tracks),
            "total_likes": sum_numbers(t["likes"] for t in tracks),
            "url": pl["url"],
        })
    releases.sort(key=lambda r: r["published_at"])
    return releases, tracks_rows


# ------------------------------ Excel-запис ------------------------------
# Опис колонок: (ключ у рядку, заголовок, ширина, формат числа)
SONG_COLUMNS = [
    ("title", "Назва", 42, None),
    ("source", "Джерело", 24, None),
    ("published_date", "Дата публікації", 16, DATE_FMT),
    ("duration", "Тривалість", 12, None),
    ("duration_sec", "Тривалість, с", 14, "0"),
    ("views", "Перегляди", 13, INT_FMT),
    ("likes", "Лайки", 10, INT_FMT),
    ("comments", "Коментарі", 12, INT_FMT),
    ("topic_categories", "Тематичні категорії", 36, None),
    ("playlists", "Плейлисти", 34, None),
    ("releases", "Релізи", 30, None),
    ("published_at", "Опубліковано (UTC)", 20, DATETIME_FMT),
    ("url", "Посилання", 46, None),
    ("video_id", "ID відео", 14, None),
]
PLAYLIST_COLUMNS = [
    ("title", "Назва", 42, None),
    ("item_count", "Відео", 10, INT_FMT),
    ("is_release", "Реліз?", 10, None),
    ("published_at", "Створено (UTC)", 20, DATETIME_FMT),
    ("url", "Посилання", 52, None),
    ("playlist_id", "ID плейлиста", 38, None),
]
RELEASE_COLUMNS = [
    ("title", "Реліз", 40, None),
    ("published_at", "Дата (UTC)", 20, DATETIME_FMT),
    ("type_guess", "Тип (припущення)", 18, None),
    ("tracks", "Треків", 9, INT_FMT),
    ("tracks_unavailable", "Недоступних треків", 19, INT_FMT),
    ("total_duration", "Сумарна тривалість", 19, None),
    ("total_views", "Сумарні перегляди", 18, INT_FMT),
    ("total_likes", "Сумарні лайки", 14, INT_FMT),
    ("url", "Посилання", 52, None),
    ("release_id", "ID релізу", 38, None),
]
RELEASE_TRACK_COLUMNS = [
    ("release_title", "Реліз", 34, None),
    ("track_no", "№", 6, "0"),
    ("title", "Трек", 40, None),
    ("duration", "Тривалість", 12, None),
    ("duration_sec", "Тривалість, с", 14, "0"),
    ("views", "Перегляди", 13, INT_FMT),
    ("likes", "Лайки", 10, INT_FMT),
    ("comments", "Коментарі", 12, INT_FMT),
    ("topic_categories", "Тематичні категорії", 36, None),
    ("in_channel_uploads", "У завантаженнях каналу", 24, None),
    ("published_date", "Дата публікації", 16, DATE_FMT),
    ("url", "Посилання", 46, None),
    ("video_id", "ID відео", 14, None),
]
COMMENT_COLUMNS = [
    ("title", "Відео", 42, None),
    ("source", "Джерело", 24, None),
    ("date", "Дата (UTC)", 14, DATE_FMT),
    ("top_level", "Коментарі верхнього рівня", 20, INT_FMT),
    ("replies", "Відповіді", 12, INT_FMT),
    ("total", "Усього", 10, INT_FMT),
    ("video_id", "ID відео", 14, None),
]

HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", start_color="1F3864")
BODY_FONT = Font(name=FONT_NAME)
LINK_FONT = Font(name=FONT_NAME, color="0563C1", underline="single")


def set_cell(cell, value):
    """Записує значення з урахуванням особливостей Excel."""
    if value == "" or value is None:
        return
    if isinstance(value, bool):
        value = "так" if value else "ні"
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        cell.value = value
        if value.startswith("="):  # щоб назва не сприймалась як формула
            cell.data_type = "s"
    else:
        cell.value = value


def write_sheet(wb, title, columns, rows):
    ws = wb.create_sheet(title)
    for c, (_, header, width, _fmt) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=c, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[1].height = 32

    for r, row in enumerate(rows, 2):
        for c, (key, _, _, fmt) in enumerate(columns, 1):
            cell = ws.cell(row=r, column=c)
            value = row.get(key, "")
            set_cell(cell, value)
            cell.font = BODY_FONT
            if fmt:
                cell.number_format = fmt
            if key == "url" and value:
                cell.hyperlink = value
                cell.font = LINK_FONT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(rows) + 1, 1)}"
    return ws


def write_info_sheet(wb, channel, counts):
    ws = wb.create_sheet("Про дані")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 95
    rows = [
        ("Канал", channel["snippet"]["title"]),
        ("ID каналу", CHANNEL_ID),
        ("Дата збору (локальний час)", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Джерело", "YouTube Data API v3"),
        ("Пісні (усього)", counts["songs"]),
        ("Плейлисти", counts["playlists"]),
        ("Релізи", counts["releases"]),
        ("", ""),
        ("Примітки", ""),
        ("Дати й час", "Усі дати та час у UTC."),
        ("Джерело пісні", f"«{SOURCE_UPLOAD}» - відео з завантажень каналу; "
                          f"«{SOURCE_RELEASE_ONLY}» - трек релізу, якого немає у завантаженнях "
                          "(зазвичай автогенероване Topic-відео)."),
        ("Коментарі по датах", "Окремо коментарі верхнього рівня та відповіді на них, а також сума. "
                               "Відповіді враховуються, якщо INCLUDE_REPLIES = True."),
        ("Лайки/коментарі", "Порожня клітинка означає, що власник приховав лічильник або вимкнув коментарі."),
        ("Тематичні категорії", "Беруться з topicDetails API (назви статей Вікіпедії)."),
        ("Тип релізу", "Це припущення: 1 трек = сингл, інакше альбом/EP."),
        ("Треки релізів", "Часто це окремі автогенеровані відео (Topic), тому їхні перегляди "
                          "можуть відрізнятися від переглядів офіційного кліпу."),
    ]
    for r, (label, value) in enumerate(rows, 1):
        a, b = ws.cell(row=r, column=1), ws.cell(row=r, column=2)
        set_cell(a, label)
        set_cell(b, value)
        a.font = Font(name=FONT_NAME, bold=True)
        b.font = BODY_FONT
        b.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        a.alignment = Alignment(vertical="top")


def save_workbook(wb):
    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        sys.exit(f"Не вдалося записати {OUTPUT_FILE}: закрийте файл у Excel і запустіть знову.")
    print(f"Збережено: {os.path.abspath(OUTPUT_FILE)}")


# ---------------------------------- main ---------------------------------
def main():
    if not API_KEY:
        sys.exit("Задайте API-ключ у змінній середовища YOUTUBE_API_KEY.")

    try:
        channel = get_channel()
        print(f"Канал: {channel['snippet']['title']}")
        stats = channel["statistics"]
        print(f"  відео: {stats.get('videoCount')}, "
              f"перегляди: {stats.get('viewCount')}, "
              f"підписники: {stats.get('subscriberCount', 'приховано')}")

        uploads_id = channel["contentDetails"]["relatedPlaylists"]["uploads"]

        print("Отримую список відео...")
        video_ids = get_all_video_ids(uploads_id)
        print(f"  знайдено: {len(video_ids)}")

        print("Отримую деталі відео...")
        videos = get_video_details(video_ids)
        details_by_id = {v["id"]: v for v in videos}

        print("Отримую плейлисти...")
        playlists = get_playlists()
        known = {p["playlist_id"] for p in playlists}
        manual = [i for i in RELEASE_PLAYLIST_IDS if i not in known]
        if manual:
            playlists += get_playlists_by_id(manual)
        n_rel = sum(p["is_release"] for p in playlists)
        print(f"  плейлистів: {len(playlists)}, з них релізів: {n_rel}")
        if n_rel == 0:
            print("  ! Релізи не знайдено серед плейлистів каналу. "
                  "Додайте ID релізів у RELEASE_PLAYLIST_IDS.")

        mapping, items_by_playlist = map_videos_to_playlists(playlists)

        # Треки релізів часто є окремими відео (авто-згенерований Topic-канал),
        # яких немає серед завантажень каналу, тому добираємо їх окремо.
        release_vids = {vid for p in playlists if p["is_release"]
                        for vid in items_by_playlist.get(p["playlist_id"], [])}
        missing = [v for v in release_vids if v not in details_by_id]
        if missing:
            print(f"Отримую деталі {len(missing)} треків релізів...")
            for v in get_video_details(missing):
                details_by_id[v["id"]] = v
    except ApiError as e:
        sys.exit(f"Помилка API: {e}")

    # --- Пісні ---
    # Спершу завантаження каналу, потім треки релізів, яких серед них немає.
    upload_set = set(video_ids)
    release_only_ids = sorted(release_vids - upload_set)
    song_rows = []
    for vid in list(dict.fromkeys(video_ids)) + release_only_ids:
        v = details_by_id.get(vid)
        if v is None:  # приватне/видалене відео
            continue
        row = base_row(v)
        if ONLY_MUSIC and "music" not in row["topic_categories"].lower():
            continue
        row["source"] = SOURCE_UPLOAD if vid in upload_set else SOURCE_RELEASE_ONLY
        m = mapping.get(vid, {})
        row["playlists"] = "; ".join(m.get("playlists", []))
        row["releases"] = "; ".join(m.get("releases", []))
        song_rows.append(row)
    song_rows.sort(key=lambda r: r["published_at"])

    # --- Релізи ---
    releases, release_tracks = build_releases(
        playlists, items_by_playlist, details_by_id, set(video_ids))

    # --- Коментарі по датах ---
    comment_rows = []
    if COLLECT_COMMENTS:
        print("Рахую коментарі по датах...")
        for i, row in enumerate(song_rows, 1):
            if row["comments"] == 0:
                continue
            print(f"  [{i}/{len(song_rows)}] {row['title']}")
            top, replies = count_comments_by_date(row["video_id"])
            for d in sorted(set(top) | set(replies)):
                comment_rows.append({
                    "video_id": row["video_id"],
                    "title": row["title"],
                    "source": row["source"],
                    "date": date.fromisoformat(d),
                    "top_level": top[d],
                    "replies": replies[d],
                    "total": top[d] + replies[d],
                })

    # --- Запис у Excel ---
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "Пісні", SONG_COLUMNS, song_rows)
    write_sheet(wb, "Плейлисти", PLAYLIST_COLUMNS, playlists)
    write_sheet(wb, "Релізи", RELEASE_COLUMNS, releases)
    write_sheet(wb, "Треки релізів", RELEASE_TRACK_COLUMNS, release_tracks)
    if COLLECT_COMMENTS:
        columns = COMMENT_COLUMNS if INCLUDE_REPLIES else [
            c for c in COMMENT_COLUMNS if c[0] not in ("replies", "total")]
        write_sheet(wb, "Коментарі по датах", columns, comment_rows)
    write_info_sheet(wb, channel, {
        "songs": len(song_rows), "playlists": len(playlists), "releases": len(releases)})
    save_workbook(wb)


if __name__ == "__main__":
    main()
