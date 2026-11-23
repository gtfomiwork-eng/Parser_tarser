import csv
import hashlib
import io
import json
import random
import re
import shutil
import time
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import quote, urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup, Comment

PROGRAM = "yandex_search_list"
VERSION = 2
PROJECT_NAME = "Парсинг Яндекс"
CHAT_NAME = "Парсинг поиска Яндекса"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/128.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9",
}
BASE = BOOKS_SITE = "https://books.yandex.ru"
DEBUG_DIR = Path("debug_html")
FILENAME_FORMAT = "Ybook_searchlist_date%Y%m%d_time%H%M%S.xlsx"

# Вкладки поиска, по которым ищем: название -> тип в адресе /search/<тип>/<запрос>
SEARCH_TABS = {"Книги": "books", "Аудиокниги": "audiobooks", "Комиксы": "comicbooks"}
RESULTS_PER_TAB = 20        # сколько первых результатов вкладки рассматривать (первая порция поиска — 20)
MAX_CARDS_PER_BOOK = 15     # предохранитель: не больше стольких карточек-кандидатов на одну книгу перечня
MAX_EXTRA_PER_BOOK = 20     # не больше стольких дополнительных версий (с карточек) на одну книгу перечня

# Пороги совпадения (0…1) — см. раздел 4
TITLE_EXACT, TITLE_LIKELY, TITLE_PARTIAL = 0.90, 0.70, 0.50

# Встроенный перечень из задачи (Автор; Название; ASTA). Можно заменить своим или дать файл при запуске.
DEFAULT_LIST = """Автор;Название;ASTA
Юрий Лотман;Роман А.С. Пушкина "Евгений Онегин";ASTA00000000066626
Ники Сью;Люблю тебя ненавидеть;ASTA00000000067986
Михаил Хорс;Бросить курить. Как? Легкий способ - минус привычка - плюс я;ASTA00000000068495
Пьердоменико Баккаларио, Алессандро Гатти;Отличная работа, Уилл Мугли!;ASTA00000000067993
Ерофей Трофимов;Толмач. Кровавая схватка;ASTA00000000067893
Полина Серебрякова;Мне не больно. Как помочь себе жить дальше после ПТСР и психологических травм и начать доверять жизни;ASTA00000000068780
Ник Тошес;Кодекс Данте;ASTA00000000068304
Дэвид Вебер;Космическая станция "Василиск";ASTA00000000068562
Юлия Каштанова;Команда Флоры. Новые приключения;ASTA00000000068505
Алекс Каменев;Хлад-3;ASTA00000000068977
Александр Койвунен;Перспектива;ASTA00000000067998
Дмитрий Кружевский;Отпуск?!;ASTA00000000068006
Карл фон Клаузевиц;О войне;ASTA00000000066691
Андрей Сарабьянов;Давид Бурлюк в России;ASTA00000000066791
Антон Безмолитвенный;Буддизм. От сутр до нейросетей;ASTA00000000067803
"""

# Метки блока «О книге» в карточке — чтобы обрезать значение на соседнем поле
KNOWN_LABELS = [
    "Возрастные ограничения", "Правообладатель", "Год выхода издания", "Год выхода",
    "Издательство", "Бумажных страниц", "Переводчики", "Переводчик", "Серия",
    "ISBN", "Длительность", "Чтецы", "Чтец", "Рассказчики", "Рассказчик", "Дата публикации",
]
# Эмоции сайта: внутреннее имя -> (значок, подпись на сайте). Порядок = порядок колонок в Excel.
EMOTIONS = {
    "thumbsup": ("👍", "Советую"), "thumbsdown": ("👎", "Не советую"),
    "rocket": ("🚀", "Не оторваться"), "droplet": ("💧", "До слёз"),
    "zzz": ("💤", "Скучно"), "bulb": ("💡", "Познавательно"), "crystal_ball": ("🔮", "Мудро"),
    "dart": ("🎯", "Полезно"), "revolving_hearts": ("💞", "Романтично"), "panda_face": ("🐼", "Мило"),
    "palm_tree": ("🌴", "В отпуск"), "smile": ("😄", "Смешно"), "skull": ("💀", "Жесть"),
    "see_no_evil": ("🙈", "Ничего не понятно"), "shit": ("💩", "Фуууу"),
}
WARNING_TAGS = {"drugs": "упоминание наркотиков"}


# ===== ячейка 6 =====
DELAY_MIN, DELAY_MAX = 2.0, 4.0      # пауза между запросами, сек (случайная в этом диапазоне)
LONG_BREAK_EVERY = 100               # после стольких запросов к сайту...
LONG_BREAK_SECONDS = 60              # ...перерыв на столько секунд
COOLDOWNS = [60, 180, 600]           # паузы после признаков ограничения, сек (1-й, 2-й, 3-й раз подряд)
CACHE_DIR = Path("cache_html")       # кэш скачанных страниц; в Colab можно указать папку на Google Диске
CACHE_HOURS = 24                     # сколько часов страница в кэше считается свежей (0 — кэш выключен)
INCOGNITO = True                     # как окно «Инкогнито»: без входа в аккаунт и без cookies между запросами

STATS = {"с сайта": 0, "из кэша": 0, "ограничений": 0, "повторов": 0}
_session = requests.Session()
_session.headers.update(HEADERS)
_state = {"last": 0.0, "slow": 1.0}


class SiteBlocked(Exception):
    """Сайт несколько раз подряд ограничил запросы — продолжать сейчас бессмысленно."""


def _cache_file(url):
    return CACHE_DIR / (hashlib.md5(url.encode("utf-8")).hexdigest() + ".html")


def _wait_turn():
    delay = random.uniform(DELAY_MIN, DELAY_MAX) * _state["slow"]
    rest = delay - (time.time() - _state["last"])
    if rest > 0:
        time.sleep(rest)
    if STATS["с сайта"] and STATS["с сайта"] % LONG_BREAK_EVERY == 0:
        print(f"\n  … плановый перерыв {LONG_BREAK_SECONDS} с после {STATS['с сайта']} запросов")
        time.sleep(LONG_BREAK_SECONDS)


def block_reason(resp):
    """Причина, по которой ответ похож на ограничение со стороны сайта, или None."""
    if resp.status_code in (403, 429):
        return f"код ответа {resp.status_code}"
    if re.search(r"showcaptcha|checkcaptcha", resp.url, re.I):
        return "капча"
    head = resp.text[:30000]
    normal = 'rel="canonical"' in head or 'property="og:url"' in head or "<title>" in head
    if not normal and re.search(r"captcha|не робот", head, re.I):
        return "капча"
    return None


def looks_complete(html):
    """Страница дошла до конца (не оборвалась на середине)."""
    return "</html>" in html[-3000:].lower()


def fetch(url, use_cache=True):
    path = _cache_file(url)
    if use_cache and CACHE_HOURS and path.exists() and time.time() - path.stat().st_mtime < CACHE_HOURS * 3600:
        STATS["из кэша"] += 1
        return path.read_text(encoding="utf-8")

    blocks, problems = 0, 0
    while True:
        _wait_turn()
        if INCOGNITO:
            _session.cookies.clear()      # каждый запрос — как из нового окна «Инкогнито»
        try:
            r = _session.get(url, timeout=30)
        except (requests.ConnectionError, requests.Timeout,
                requests.exceptions.ChunkedEncodingError, requests.exceptions.ContentDecodingError) as e:
            problems += 1                 # нет связи, тайм-аут или ответ оборвался на середине — повторяем
            STATS["повторов"] += 1
            if problems >= 4:
                raise
            print(f"\n  ⚠ нет связи ({type(e).__name__}), повтор через {20 * problems} с")
            time.sleep(20 * problems)
            continue
        finally:
            _state["last"] = time.time()
        STATS["с сайта"] += 1

        reason = block_reason(r)
        if reason:
            STATS["ограничений"] += 1
            _state["slow"] = min(_state["slow"] * 1.5, 4)        # дальше идём медленнее
            if blocks >= len(COOLDOWNS):
                raise SiteBlocked(f"сайт ограничил запросы ({reason}) {blocks + 1} раза подряд")
            wait = COOLDOWNS[blocks]
            retry_after = r.headers.get("Retry-After", "")
            if retry_after.isdigit():
                wait = max(wait, int(retry_after))
            print(f"\n  ⚠ похоже, сайт ограничил запросы ({reason}). Пауза {wait // 60} мин {wait % 60} с, "
                  f"дальше паузы между запросами ×{_state['slow']:.1f}")
            time.sleep(wait)
            blocks += 1
            continue

        if r.status_code >= 250:
            problems += 1
            STATS["повторов"] += 1
            if problems >= 3:
                r.raise_for_status()
            time.sleep(20 * problems)
            continue
        r.raise_for_status()                                       # 404 и т.п. — сразу ошибка

        if not looks_complete(r.text):                             # страница пришла обрезанной
            problems += 1
            STATS["повторов"] += 1
            if problems >= 3:
                raise ValueError("страница приходит неполной (возможно, ограничение сайта)")
            print(f"\n  ⚠ страница пришла неполной, повтор через {30 * problems} с")
            time.sleep(30 * problems)
            continue

        if CACHE_HOURS:
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            path.write_text(r.text, encoding="utf-8")
        return r.text


def net_summary():
    return ", ".join(f"{k}: {v}" for k, v in STATS.items())


def ask_clear_cache():
    """При запуске: очистить кэш (по умолчанию) или оставить, чтобы продолжить прерванный прогон."""
    files = list(CACHE_DIR.glob("*.html")) if CACHE_DIR.exists() else []
    if not files:
        return
    age_h = (time.time() - min(f.stat().st_mtime for f in files)) / 3600
    ans = input(f"В кэше {len(files)} страниц (самой старой {age_h:.0f} ч). Очистить кэш и скачать всё заново? "
                f"(Enter — да; «н» — нет, продолжить прерванный прогон): ").strip().lower()
    if ans in ("н", "нет", "n", "no"):
        print("Кэш оставлен: уже скачанные страницы возьмутся из него (данные в них — на момент скачивания).")
        return
    shutil.rmtree(CACHE_DIR, ignore_errors=True)
    print(f"Кэш очищен ({len(files)} страниц) — все страницы будут скачаны заново.")


# ===== ячейка 8 =====
ASTA_RE = re.compile(r"ASTA\d+", re.I)


def norm(s):
    s = str(s or "").lower().replace("ё", "е")
    s = re.sub(r"[^\w]+", " ", s)
    return re.sub(r"\s+", " ", s.replace("_", " ")).strip()


def surnames(authors):
    """'Пьердоменико Баккаларио, Алессандро Гатти' -> ['баккаларио', 'гатти']; 'Карл фон Клаузевиц' -> ['клаузевиц']."""
    out = []
    for a in re.split(r"[,;]| и ", authors or ""):
        words = [w for w in norm(a).split() if len(w) > 1]
        if words:
            out.append(words[-1])
    return out


def _surname_in(sur, text_words):
    stem = sur if len(sur) <= 4 else sur[:max(4, len(sur) - 2)]
    return any(w == sur or (len(sur) > 4 and w.startswith(stem)) for w in text_words)


def author_match(want_authors, got_text):
    words = norm(got_text).split()
    return any(_surname_in(s, words) for s in surnames(want_authors))


def title_score(want, got, long_text=False):
    """Похожесть названия из перечня и названия (или текста карточки результата) на сайте, 0…1."""
    a, b = norm(want), norm(got)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if long_text:                                   # описание результата поиска: ищем название внутри текста
        if f" {a} " in f" {b} ":
            return 0.95
    elif (b.startswith(a + " ") or a.startswith(b + " ")) and min(len(a), len(b)) >= 4:
        return 0.9
    ta = [w for w in a.split() if len(w) >= 3] or a.split()
    tb = b.split()
    cover = sum(any(x == w or (len(w) > 4 and x.startswith(w[:-1])) for x in tb) for w in ta) / len(ta)
    ratio = 0.0 if long_text else SequenceMatcher(None, a, b).ratio()
    return round(max(ratio, cover * (0.85 if cover == 1 else 0.75)), 3)


def title_score_n(want, got, long_text=False):
    """title_score + номера: «Хлад-3» и «Хлад-2», «Отпуск» и «Отпуск 2» — разные книги (не выше «проверить»)."""
    t = title_score(want, got, long_text)
    dw, dg = set(re.findall(r"\d+", norm(want))), set(re.findall(r"\d+", norm(got)))
    if not long_text and dw != dg:
        t = min(t, 0.6)
    elif long_text and dw and not dw <= dg:
        t = min(t, 0.6)
    return t


def match_level(t, author_ok):
    if t >= TITLE_EXACT and author_ok:
        return "точно"
    if t >= TITLE_LIKELY and author_ok:
        return "вероятно"
    if (t >= TITLE_EXACT and not author_ok) or (t >= TITLE_PARTIAL and author_ok):
        return "проверить"
    return "не подходит"


LEVEL_ORDER = {"точно": 0, "вероятно": 1, "проверить": 2, "не подходит": 3}


def title_main(title):
    """Начало названия до первой точки/вопроса/двоеточия (не режем «А.С. Пушкина» на инициалах)."""
    for m in re.finditer(r"[.?!:…]+(?=\s|$)", title):
        last = re.split(r"[.\s]", title[:m.start()].strip())[-1].strip("«»\"'")
        if len(last) > 1:                                  # «А.С.» — инициалы, здесь не режем
            head = title[:m.start()].strip(" «»\"'")
            return head if head and head != title.strip() else title
    return title


def clean_query(s):
    return re.sub(r"\s+", " ", re.sub(r"[«»\"“”„']", " ", s)).strip()


def queries_for(item):
    """Варианты запроса по порядку: «Фамилия Название», «Фамилия Начало названия», «Название»."""
    sur = (item["author"] or "").split(",")[0].split()
    sur = sur[-1] if sur else ""
    full, main = clean_query(item["title"]), clean_query(title_main(item["title"]))
    out = []
    for q in (f"{sur} {full}", f"{sur} {main}", full):
        q = q.strip()
        if q and q not in out:
            out.append(q)
    return out


def read_text_any(path):
    """Текст CSV/TXT в любой кодировке или .xlsx (первые листы склеиваются через ;)."""
    raw = Path(path).read_bytes()
    if raw[:2] == b"PK":
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        return "\n".join(";".join("" if v is None else str(v) for v in row)
                         for ws in wb.worksheets for row in ws.iter_rows(values_only=True))
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff") or raw[:200].count(b"\x00") > 20:
        return raw.decode("utf-16", errors="replace")
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


def parse_list(text):
    """Перечень из текста с разделителем ; (или табуляцией / запятой). Колонки ищутся по заголовку
    (Автор / Название / ASTA), без заголовка — ASTA по виду кода, остальные: 1-я — автор, 2-я — название."""
    sample = text[:2000]
    delim = ";" if ";" in sample else ("\t" if "\t" in sample else ",")
    rows = [r for r in csv.reader(io.StringIO(text), delimiter=delim) if any(c.strip() for c in r)]
    if not rows:
        return []
    head = [norm(c) for c in rows[0]]
    col = {}
    for i, h in enumerate(head):
        if "автор" in h and "author" not in col:
            col["author"] = i
        elif ("назван" in h or "заглав" in h) and "title" not in col:
            col["title"] = i
        elif "asta" in h and "asta" not in col:
            col["asta"] = i
    if {"author", "title"} <= set(col):
        rows = rows[1:]
    items = []
    for r in rows:
        r = [c.strip() for c in r]
        if col.get("author") is not None and col.get("title") is not None:
            author = r[col["author"]] if col["author"] < len(r) else ""
            title = r[col["title"]] if col["title"] < len(r) else ""
            asta = r[col["asta"]] if col.get("asta") is not None and col["asta"] < len(r) else ""
        else:
            asta = next((c for c in r if ASTA_RE.fullmatch(c)), "")
            rest = [c for c in r if c and c != asta]
            author, title = (rest + ["", ""])[:2]
        if not asta:
            m = ASTA_RE.search(";".join(r))
            asta = m.group(0) if m else ""
        if title:
            items.append({"n": len(items) + 1, "author": author, "title": title, "asta": asta})
    return items


# ===== ячейка 10 =====
def parse_count(s):
    """'7.7K' -> 7700, '175.8K' -> 175800, '1.2M' -> 1200000, '6 721' -> 6721."""
    if not s:
        return None
    s = s.replace(" ", "").replace(" ", "").replace(",", ".").upper()
    m = re.search(r"(\d+(?:\.\d+)?)([KMКМ]?)", s)
    if not m:
        return None
    mult = {"": 1, "K": 1_000, "К": 1_000, "M": 1_000_000, "М": 1_000_000}[m.group(2)]
    return int(round(float(m.group(1)) * mult))


def meta(soup, prop):
    tag = soup.find("meta", attrs={"property": prop}) or soup.find("meta", attrs={"name": prop})
    return tag["content"].strip() if tag and tag.get("content") else None


def text_of(el):
    """Текст элемента без HTML-комментариев и скриптов (React вставляет <!-- --> между кусками текста)."""
    text = " ".join(
        s.strip() for s in el.find_all(string=True)
        if not isinstance(s, Comment) and s.strip() and s.parent.name not in ("script", "style", "noscript")
    )
    return re.sub(r"[    ]", " ", text)


def tab_counter(soup, suffix):
    a = soup.find("a", href=re.compile(rf"/(?:books|audiobooks|comicbooks|serials)/[^/?#]+/{suffix}/?(?:[?#].*)?$"))
    if not a:
        return None, None
    raw = re.sub(r"^[^\d]+", "", text_of(a).replace(" ", ""))
    return (raw or None), parse_count(raw)


def exact_section_counter(soup, suffix):
    for h2 in soup.find_all("h2"):
        a = h2.find("a", href=re.compile(rf"/{suffix}/?(?:[?#].*)?$"))
        if a:
            digits = re.sub(r"\D", "", text_of(a))
            return int(digits) if digits else None
    return None


def _cut_at_next_label(value, label):
    others = [l for l in KNOWN_LABELS if l != label]
    positions = [m.start() for l in others for m in [re.search(re.escape(l), value)] if m]
    return value[:min(positions)].strip(" :;,") if positions else value.strip(" :;,")


def labeled_value(soup, label):
    """Значение после метки 'Правообладатель: АСТ'. Возвращает (текст, ссылка или None)."""
    lab = re.escape(label)
    label_node = re.compile(rf"^\s*{lab}\s*(?::.*)?$", re.S)
    for node in soup.find_all(string=label_node):
        if isinstance(node, Comment) or node.parent.name in ("script", "style"):
            continue
        el = node.parent
        for _ in range(4):
            if el is None:
                break
            text = text_of(el)
            m = re.search(rf"(?:^|\s){lab}\s*:?\s*(.*)$", text, re.S)
            if not m:
                break
            value = _cut_at_next_label(m.group(1), label)
            if value:
                link = None
                for a in el.find_all("a", href=True):
                    if text_of(a) and text_of(a) in value:
                        link = a["href"]
                        break
                return value, link
            el = el.parent
    page = text_of(soup.body or soup)
    m = re.search(rf"{lab}\s*:\s*(.+)", page)
    if m:
        value = _cut_at_next_label(m.group(1)[:200], label)
        return (value or None), None
    return None, None


EDITION_RE = re.compile(r"/(books|audiobooks|comicbooks|serials)/([A-Za-z0-9]+)/?(?:[?#].*)?$")
TYPE_NAMES = {"books": "текст", "audiobooks": "аудио", "comicbooks": "комикс", "serials": "сериал"}
FORMAT_NAMES = {"books": ("читать", "электронная книга"), "audiobooks": ("слушать", "аудиокнига"),
                "comicbooks": ("читать", "комикс"), "serials": ("читать", "сериал")}


def edition_of(href):
    m = EDITION_RE.search(href or "")
    return (m.group(1), m.group(2)) if m else (None, None)


def meta_link(soup, rel):
    tag = soup.find("link", rel=rel)
    return tag["href"] if tag and tag.get("href") else None


def other_versions(soup, own_id):
    result = {"versions_shown": None, "versions_total": 0, "versions_text": 0,
              "versions_audio": 0, "versions_comics": 0, "versions_urls": []}
    head = None
    for h in soup.find_all(["h2", "h3"]):
        if re.match(r"\s*Другие версии", text_of(h), re.I):
            head = h
            break
    if head is None:
        return result
    digits = re.sub(r"\D", "", text_of(head))
    result["versions_shown"] = int(digits) if digits else None
    box = head.parent
    for _ in range(5):
        if box is None:
            break
        links = [a["href"] for a in box.find_all("a", href=EDITION_RE)]
        if any(edition_of(h)[1] != own_id for h in links):
            break
        box = box.parent
    if box is None:
        return result
    seen = []
    for a in box.find_all("a", href=EDITION_RE):
        t, i = edition_of(a["href"])
        key = f"{t}/{i}"
        if i == own_id or key in seen:
            continue
        seen.append(key)
    result["versions_urls"] = [f"{BOOKS_SITE}/{k}" for k in seen]
    result["versions_total"] = len(seen)
    result["versions_text"] = sum(k.startswith("books/") for k in seen)
    result["versions_audio"] = sum(k.startswith("audiobooks/") for k in seen)
    result["versions_comics"] = sum(k.startswith("comicbooks/") for k in seen)
    return result


NEXT_PUSH_RE = re.compile(r"^\s*self\.__next_f\.push\((\[.*\])\)\s*;?\s*$", re.S)
NEXT_KEY = '{"type":"next","value":'
_json = json.JSONDecoder()


def next_values(soup):
    """Все ответы сервера, вшитые в страницу (__next_f и ApolloSSRDataTransport): список {'data': {...}}."""
    parts, out = [], []
    for sc in soup.find_all("script"):
        s = sc.string or ""
        m = NEXT_PUSH_RE.match(s)
        if m:
            try:
                arr = json.loads(m.group(1))
            except ValueError:
                continue
            if len(arr) > 1 and isinstance(arr[1], str):
                parts.append(arr[1])
        elif "ApolloSSRDataTransport" in s:
            parts.append(s)                              # JSON прямо в тексте скрипта
    stream = "\n".join(parts)
    i = 0
    while True:
        i = stream.find(NEXT_KEY, i)
        if i < 0:
            break
        try:
            obj, end = _json.raw_decode(stream, i)
        except ValueError:
            i += len(NEXT_KEY)
            continue
        if isinstance(obj.get("value"), dict):
            out.append(obj["value"])
        i = end
    return out


def _merge(dst, src):
    for k, v in src.items():
        if v is not None or k not in dst:
            dst[k] = v


def book_data(soup, own_id):
    res = {"book": {}, "root": {}, "old": {}, "emotions": None, "found": False}
    for value in next_values(soup):
        data = value.get("data")
        if not isinstance(data, dict):
            continue
        for root, r in data.items():
            if not isinstance(r, dict):
                continue
            if root in ("textBook", "audioBook", "comicBook"):
                b = r.get("book")
                if isinstance(b, dict) and b.get("uuid") == own_id:
                    res["found"] = True
                    _merge(res["book"], b)
                    _merge(res["root"], {k: v for k, v in r.items() if k not in ("book", "headers", "__typename")})
            elif root.startswith("old"):
                for v in r.values():
                    if isinstance(v, dict) and v.get("uuid") == own_id:
                        _merge(res["old"], v)
            elif root == "emotionRating" and isinstance(r.get("emotion_ratings"), list):
                res["emotions"] = r["emotion_ratings"]
    return res


def _find_offers(node):
    if isinstance(node, dict):
        if "offers" in node:
            return node["offers"]
        for v in node.values():
            found = _find_offers(v)
            if found:
                return found
    elif isinstance(node, list):
        for v in node:
            found = _find_offers(v)
            if found:
                return found
    return None


def ld_offer(soup):
    for sc in soup.find_all("script", type="application/ld+json"):
        try:
            offers = _find_offers(json.loads(sc.string or ""))
        except ValueError:
            continue
        if isinstance(offers, list):
            offers = offers[0] if offers else None
        if isinstance(offers, dict):
            return offers
    return None


def sale_fields(soup, page_text, own_id, own_type=None):
    d = book_data(soup, own_id)
    book, old = d["book"], d["old"]
    offer = book.get("purchaseOffer") if "purchaseOffer" in book else None
    ld = ld_offer(soup)
    buy = re.search(r"Купить[^₽]{0,40}?\d[\d  ]*\s*₽", page_text)
    if offer:
        sub_from = offer.get("availableInSubscriptionFrom")
        status = f"Продажа, в подписке с {sub_from}" if sub_from else "Продажа (PPD)"
    elif "purchaseOffer" in book:
        status = "Подписка"
    elif ld or buy:
        status = "Продажа (PPD)"
    else:
        status = None if not d["found"] else "Подписка"
    price = (offer or {}).get("price")
    if price is None and ld:
        price = ld.get("price")
    if price is None and buy:
        price = parse_count(re.sub(r"[^\d]", "", buy.group(0)))
    warnings = []
    for t in book.get("tags") or []:
        if isinstance(t, dict) and str(t.get("value")).lower() == "true":
            warnings.append(WARNING_TAGS.get(t.get("name"), t.get("name")))
    emotions = {}
    for e in d["emotions"] or []:
        name = (e.get("emotion") or {}).get("name")
        if name:
            emotions[name] = e.get("count")
    auditory = (book.get("auditory") or {}).get("total")
    return {
        "sale_status": status, "price": price,
        "currency": (offer or {}).get("currency") or (ld or {}).get("priceCurrency"),
        "in_subscription_from": (offer or {}).get("availableInSubscriptionFrom"),
        "buy_button": re.sub(r"\s+", " ", buy.group(0)) if buy else None,
        "price_shown": "да" if buy else ("нет" if offer else None),
        "ld_price": (ld or {}).get("price"),
        "ld_availability": ((ld or {}).get("availability") or "").rsplit("/", 1)[-1] or None,
        "free_fragment": "да" if re.search(r"Бесплатный отрывок", page_text) else "нет",
        "web_reading": "только в приложении" if re.search(r"Читать в приложении", page_text) else None,
        "warnings": "; ".join(warnings) or None,
        "readers_exact": old.get("readers_count") if old.get("readers_count") is not None else
                         (auditory if own_type != "audiobooks" else None),
        "listeners_exact": old.get("listeners_count") if old.get("listeners_count") is not None else
                           (auditory if own_type == "audiobooks" else None),
        "impressions_data": old.get("impressions_count"),
        "quotes": old.get("quotes_count"), "on_shelves": old.get("bookshelves_count"),
        "paper_pages_data": old.get("paper_pages") or old.get("pages_count"),
        "emotions": emotions, "data_found": d["found"],
    }


def parse_book(html, url=None):
    soup = BeautifulSoup(html, "lxml")
    page_text = text_of(soup.body or soup)

    author = meta(soup, "og:book:author")
    h1 = soup.find("h1")
    author_link = None
    if h1:
        a = h1.find("a", href=re.compile(r"/authors/")) or h1.parent.find("a", href=re.compile(r"/authors/"))
        if a:
            author_link = a["href"]
            author = author or text_of(a)
    title = None
    if h1:
        title = " ".join(s.strip() for s in h1.find_all(string=True, recursive=False)
                         if not isinstance(s, Comment) and s.strip())
        if not title:
            title = text_of(h1)
            if author:
                title = title.replace(author, "").strip()
    if not title:
        og = (meta(soup, "og:title") or "").replace(" ", " ")
        title = og.split("—")[0].replace(author or "", "").strip() or None

    own_type, own_id = edition_of(meta_link(soup, "canonical") or meta(soup, "og:url") or url or "")
    if not own_id:
        own_type, own_id = edition_of(url or "")
    sale = sale_fields(soup, page_text, own_id, own_type)

    def emotion(word, name):
        pat = r"\s+".join(re.escape(w) for w in word.split()).replace("е", "[её]")
        m = re.search(rf"(?<![А-Яа-яЁё])(?<![Нн]е\s){pat}\s*(\d+)(?![\d.,]*[KКMМ]\b)", page_text, re.I)
        return int(m.group(1)) if m else sale["emotions"].get(name)

    impressions_raw, impressions_approx = tab_counter(soup, "impressions")
    readers_raw, readers_approx = tab_counter(soup, "readers")
    listeners_raw, listeners_approx = tab_counter(soup, "listeners")

    rightholder, _ = labeled_value(soup, "Правообладатель")
    publisher, publisher_link = labeled_value(soup, "Издательство")
    narrator = labeled_value(soup, "Рассказчик")[0] or labeled_value(soup, "Чтец")[0]
    translator = labeled_value(soup, "Переводчик")[0] or labeled_value(soup, "Переводчики")[0]
    duration = labeled_value(soup, "Длительность")[0]
    year = labeled_value(soup, "Год выхода издания")[0]
    pages = labeled_value(soup, "Бумажных страниц")[0]
    age = labeled_value(soup, "Возрастные ограничения")[0]

    pair_url = None
    for a in soup.find_all("a", href=True):
        if text_of(a) in ("Текст", "Аудио", "Комикс"):
            t, i = edition_of(a["href"])
            if i and i != own_id:
                pair_url = f"{BOOKS_SITE}/{t}/{i}"
                break
    formats = []
    pair_type = edition_of(pair_url)[0] if pair_url else None
    for t, link in ((own_type, url), (pair_type, pair_url)):
        if t in FORMAT_NAMES and t not in [f["type"] for f in formats]:
            action, name = FORMAT_NAMES[t]
            formats.append({"type": t, "action": action, "format": name, "url": link})

    versions = other_versions(soup, own_id)
    tag_names = [t.strip() for t in (meta(soup, "og:book:tag") or "").split(",") if t.strip()]
    tags, seen = [], set()
    for a in soup.find_all("a", href=re.compile(r"/topic/all/")):
        if a.find_parent(["ol", "nav"]):
            continue
        name = text_of(a)
        if name in seen or (tag_names and name not in tag_names):
            continue
        seen.add(name)
        tags.append({"name": name, "url": a["href"]})
    if not tags:
        tags = [{"name": n, "url": None} for n in tag_names]

    return {
        "url": url, "type": TYPE_NAMES.get(own_type),
        "edition_id": f"{own_type}/{own_id}" if own_id else None,
        **{k: sale[k] for k in ("sale_status", "price", "currency", "in_subscription_from", "buy_button",
                                "price_shown", "ld_price", "ld_availability", "free_fragment",
                                "web_reading", "warnings")},
        "pair_url": pair_url, "author": author, "author_url": author_link, "title": title,
        "impressions": impressions_raw,
        "impressions_exact": exact_section_counter(soup, "impressions") or sale["impressions_data"] or impressions_approx,
        "recommend": emotion("Советую", "thumbsup"), "not_recommend": emotion("Не советую", "thumbsdown"),
        "cant_put_down": emotion("Не оторваться", "rocket"), "to_tears": emotion("До слез", "droplet"),
        "readers": readers_raw, "readers_approx": readers_approx, "readers_exact": sale["readers_exact"],
        "listeners": listeners_raw, "listeners_approx": listeners_approx, "listeners_exact": sale["listeners_exact"],
        "quotes": sale["quotes"], "on_shelves": sale["on_shelves"],
        "rightholder": rightholder, "publisher": publisher, "publisher_url": publisher_link,
        "narrator": narrator, "translator": translator, "duration": duration, "year": year,
        "pages": parse_count(pages) if pages else sale["paper_pages_data"], "age": age,
        "formats": formats, "tags": tags, **versions, "emotions": sale["emotions"],
    }


def publisher_group(rightholder, publisher):
    """Наши (АСТ, Азбука и импринты) / ЭКСМО / МИФ / Конкуренты — по полям «Правообладатель» и «Издательство»."""
    fields = [f for f in (rightholder, publisher) if f]
    if not fields:
        return "нет данных"
    for f in fields:
        if re.search(r"аст[\s-]*пресс", f, re.I):
            continue                                            # старое «Аст-пресс» — не наше
        if re.search(r"(?<![А-ЯA-Z])АСТ(?![А-ЯA-Z])", f) or re.search(r"азбука|аттикус|астрель|махаон", f, re.I):
            return "Наши"
    for f in fields:
        if re.search(r"эксмо|бомбора|inspiria|freedom|like book|fanzon", f, re.I):
            return "ЭКСМО"
    for f in fields:
        if re.search(r"манн|иванов и фербер|(?<![А-Яа-я])миф(?![А-Яа-я])", f, re.I):
            return "МИФ"
    return "Конкуренты"


# ===== ячейка 12 =====
NAV_SEGMENTS = {"search", "catalog", "section", "showcase", "topic", "login", "subscription", "about",
                "help", "legal", "promo", "plus"}
EMPTY_RE = re.compile(r"ничего не (?:найдено|нашлось)|нет результатов|не нашли", re.I)
TAB_KIND = {"books": "books", "audiobooks": "audiobooks", "comicbooks": "comicbooks"}


def clean_url(href):
    p = urlsplit(urljoin(BASE, href))
    return f"{p.scheme}://{p.netloc}{p.path.rstrip('/')}"


def search_url(slug, query):
    return f"{BASE}/search/{slug}/{quote(query, safe='')}"


def embedded_titles(soup):
    """uuid -> {'name', 'authors'} из встроенных данных страницы (если они там есть)."""
    out = {}

    def walk(node):
        if isinstance(node, dict):
            uid, name = node.get("uuid"), node.get("name") or node.get("title")
            if isinstance(uid, str) and isinstance(name, str):
                names = []
                for k, v in node.items():
                    if "author" in k.lower() or k in ("persons", "people"):
                        for p in (v if isinstance(v, list) else [v]):
                            if isinstance(p, dict):
                                p = p.get("person") if isinstance(p.get("person"), dict) else p
                                if isinstance(p.get("name"), str):
                                    names.append(p["name"])
                            elif isinstance(p, str):
                                names.append(p)
                rec = out.setdefault(uid, {"name": name, "authors": []})
                rec["authors"] += [n for n in names if n not in rec["authors"]]
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)
    for value in next_values(soup):
        walk(value)
    return out


def search_results(html, slug):
    """Результаты вкладки по порядку: [{'url','kind','name','desc','authors'}], пометка «пусто на сайте»."""
    soup = BeautifulSoup(html, "lxml")
    emb = embedded_titles(soup)
    root = soup.find("main") or soup.body or soup
    for junk in root.find_all(["header", "footer", "nav"]):
        junk.decompose()
    want = TAB_KIND.get(slug)
    anchors = []
    for a in root.find_all("a", href=True):
        u = clean_url(a["href"])
        parts = [x for x in urlsplit(u).path.split("/") if x]
        if urlsplit(u).netloc != "books.yandex.ru" or len(parts) != 2 or parts[0] in NAV_SEGMENTS:
            continue
        if parts[0] == want:
            anchors.append((u, a))
    items, seen = [], set()
    for u, a in anchors:
        if u in seen:
            continue
        seen.add(u)
        same = [x for x in anchors if x[0] == u]
        name = next((text_of(x[1]) for x in same if text_of(x[1])), "")
        if not name:
            img = a.find("img", alt=True)
            name = img["alt"] if img else ""
        el, desc = a, text_of(a)
        for _ in range(4):                               # текст карточки результата: поднимаемся, пока нет соседей
            par = el.parent
            if par is None:
                break
            others = {clean_url(x["href"]) for x in par.find_all("a", href=EDITION_RE)}
            if len(others - {u}) > 0:
                break
            el, desc = par, text_of(par)
        e = emb.get(u.rsplit("/", 1)[-1], {})
        items.append({"url": u, "kind": want, "name": e.get("name") or name, "desc": desc[:300],
                      "authors": ", ".join(e.get("authors") or [])})
    return items[:RESULTS_PER_TAB], bool(EMPTY_RE.search(text_of(root)))


def score_search_item(item, it):
    """Оценка результата поиска до открытия карточки."""
    t = max(title_score_n(item["title"], it["name"]), title_score_n(item["title"], it["desc"], long_text=True))
    a_ok = author_match(item["author"], " ".join([it["authors"], it["desc"]]))
    return t, a_ok


def worth_card(t, a_ok):
    return t >= TITLE_LIKELY or (a_ok and t >= TITLE_PARTIAL)


# ===== ячейка 14 =====
CARD_COLUMNS = [
    ("url", "Ссылка"), ("edition_id", "ID издания"), ("type", "Тип"),
    ("sale_status", "Статус продажи"), ("price", "Цена"), ("currency", "Валюта"),
    ("in_subscription_from", "В подписке с"), ("buy_button", "Кнопка покупки (на сайте)"),
    ("price_shown", "Цена видна на сайте"), ("ld_price", "Цена в разметке Schema.org"),
    ("ld_availability", "Наличие (Schema.org)"), ("free_fragment", "Бесплатный отрывок"),
    ("web_reading", "Чтение на сайте"), ("warnings", "Предупреждения"),
    ("pair_url", "Парная версия (текст/аудио)"), ("author", "Автор"), ("title", "Название"),
    ("impressions", "Впечатления (на сайте)"), ("impressions_exact", "Впечатления (число)"),
    ("recommend", "Советую"), ("not_recommend", "Не советую"), ("cant_put_down", "Не оторваться"),
    ("to_tears", "До слёз"), ("readers", "Читают (на сайте)"), ("readers_approx", "Читают (≈число)"),
    ("readers_exact", "Читают (точно)"), ("listeners", "Слушают (на сайте)"),
    ("listeners_approx", "Слушают (≈число)"), ("listeners_exact", "Слушают (точно)"),
    ("quotes", "Цитаты"), ("on_shelves", "На полках"),
    ("rightholder", "Правообладатель"), ("publisher", "Издательство"), ("narrator", "Рассказчик"),
    ("translator", "Переводчик"), ("duration", "Длительность"), ("year", "Год выхода издания"),
    ("pages", "Бумажных страниц"), ("age", "Возраст"), ("formats", "Форматы"), ("tags", "Теги"),
    ("versions_shown", "Других версий (по заголовку)"), ("versions_total", "Других версий (найдено)"),
    ("versions_text", "из них текст"), ("versions_audio", "из них аудио"),
    ("versions_comics", "из них комиксы"), ("versions_urls", "Другие версии (ссылки)"),
    ("emotions_all", "Эмоции (все)"), ("error", "Ошибка"),
]
LEAD_COLUMNS = [("n", "№ в перечне"), ("asta", "ASTA"), ("list_author", "Автор (перечень)"),
                ("list_title", "Название (перечень)"), ("level", "Совпадение"), ("title_score", "Похожесть названия"),
                ("author_ok", "Автор совпал"), ("how", "Как найдено"), ("group", "Группа издательства")]
INT_COLUMNS = ["Цена", "Цена в разметке Schema.org", "Впечатления (число)", "Советую", "Не советую",
               "Не оторваться", "До слёз", "Читают (≈число)", "Читают (точно)", "Слушают (≈число)",
               "Слушают (точно)", "Цитаты", "На полках", "Бумажных страниц", "Других версий (по заголовку)",
               "Других версий (найдено)", "из них текст", "из них аудио", "из них комиксы", "№ в перечне"]
WIDTHS = {"Ссылка": 42, "ID издания": 20, "ASTA": 22, "Автор (перечень)": 22, "Название (перечень)": 34,
          "Совпадение": 12, "Как найдено": 40, "Группа издательства": 14, "Статус продажи": 18,
          "Кнопка покупки (на сайте)": 24, "Предупреждения": 22, "Парная версия (текст/аудио)": 42,
          "Автор": 22, "Название": 34, "Правообладатель": 18, "Издательство": 18, "Рассказчик": 20,
          "Переводчик": 20, "Форматы": 34, "Теги": 50, "Другие версии (ссылки)": 60, "Эмоции (все)": 40,
          "Ошибка": 30, "Статус поиска": 26, "Запросы": 50, "Ссылки найденных": 70, "Описание": 60,
          "Запрос": 40, "Название (поиск)": 34, "Итог": 14, "Параметр": 34, "Значение": 70}


def emotion_title(name):
    icon, label = EMOTIONS.get(name, ("", name))
    return f"Эмоция: {icon} {label}".replace("  ", " ")


def to_int_column(series):
    num = pd.to_numeric(series, errors="coerce")
    vals = num.dropna()
    if len(vals) and ((vals % 1 != 0).any() or (vals.abs() > 2**53).any()):
        return num
    try:
        return num.astype("Int64")
    except (TypeError, ValueError):
        return num


def style_sheet(ws, link_cols=(), freeze="A2"):
    from openpyxl.styles import Alignment, Font
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = WIDTHS.get(h, 12)
    for h in link_cols:
        if h in headers:
            col = headers.index(h)
            for row in ws.iter_rows(min_row=2):
                v = row[col].value
                if v and str(v).startswith("http"):
                    row[col].hyperlink = v
                    row[col].style = "Hyperlink"


def to_excel(list_rows, book_rows, cand_rows, about, out_dir="."):
    seen_em = {n for b in book_rows for n in (b.get("emotions") or {})}
    em_names = [n for n in EMOTIONS if n in seen_em] + sorted(seen_em - set(EMOTIONS))
    columns = LEAD_COLUMNS + CARD_COLUMNS[:-1] + [(f"em:{n}", emotion_title(n)) for n in em_names] + CARD_COLUMNS[-1:]
    flat = []
    for b in book_rows:
        r = dict(b)
        r["formats"] = "; ".join(f"{f['action']} — {f['format']}" for f in b.get("formats") or [])
        r["tags"] = "; ".join(t["name"] for t in b.get("tags") or [])
        r["versions_urls"] = "; ".join(b.get("versions_urls") or [])
        r["author_ok"] = "да" if b.get("author_ok") else "нет"
        em = b.get("emotions") or {}
        r["emotions_all"] = "; ".join(f"{EMOTIONS.get(n, ('', n))[1]} {em[n]}" for n in em_names if n in em) or None
        for n in em_names:
            r[f"em:{n}"] = em.get(n)
        flat.append({title: r.get(key) for key, title in columns})
    df = pd.DataFrame(flat, columns=[t for _, t in columns])
    for col in INT_COLUMNS + [emotion_title(n) for n in em_names]:
        if col in df:
            df[col] = to_int_column(df[col])

    path = Path(out_dir) / datetime.now().strftime(FILENAME_FORMAT)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(list_rows).to_excel(xw, index=False, sheet_name="Перечень")
        style_sheet(xw.sheets["Перечень"], freeze="E2")
        df.to_excel(xw, index=False, sheet_name="Книги")
        style_sheet(xw.sheets["Книги"], ["Ссылка", "Парная версия (текст/аудио)"], freeze="F2")
        pd.DataFrame(cand_rows).to_excel(xw, index=False, sheet_name="Кандидаты")
        style_sheet(xw.sheets["Кандидаты"], ["Ссылка"])
        pd.DataFrame(about, columns=["Параметр", "Значение"]).to_excel(xw, index=False, sheet_name="О файле")
        style_sheet(xw.sheets["О файле"])
    return path, df


# ===== ячейка 16 =====
def ask_yes(prompt, default=True):
    ans = input(f"{prompt} (Enter — {'да' if default else 'нет'}; «{'н' if default else 'д'}» — "
                f"{'нет' if default else 'да'}): ").strip().lower()
    if not ans:
        return default
    return ans in ("д", "да", "y", "yes")


def load_list():
    print("Откуда взять перечень книг (Автор; Название; ASTA)?")
    print("  1 — встроенный перечень (15 книг из задачи)")
    print("  2 — CSV/XLSX-файл (колонки Автор, Название, ASTA; разделитель CSV — «;»)")
    choice = input("Ваш выбор (Enter — 1): ").strip() or "1"
    if choice == "2":
        while True:
            path = input("Путь к файлу: ").strip().strip('"\'')
            if not Path(path).is_file():
                print("Файл не найден. Попробуйте ещё раз.")
                continue
            items = parse_list(read_text_any(path))
            if items:
                return items, f"файл {Path(path).name}"
            print("В файле не нашлось строк с названием. Проверьте колонки Автор / Название / ASTA.")
    return parse_list(DEFAULT_LIST), "встроенный перечень"


def get_card(url, cards):
    if url not in cards:
        try:
            b = parse_book(fetch(url), url)
            b["error"] = None
        except SiteBlocked:
            raise
        except Exception as e:
            b = {"url": url, "error": f"{type(e).__name__}: {e}"}
        cards[url] = b
    return cards[url]


def judge_card(item, b):
    t = title_score_n(item["title"], b.get("title") or "")
    a_ok = author_match(item["author"], b.get("author") or "")
    return t, a_ok, match_level(t, a_ok)


def process_item(item, cards, cand_rows, want_extra):
    """Поиск одной книги перечня. Возвращает (найденные издания [{url, how, t, a_ok, level}], запросы, пометки)."""
    found, used_q, notes = {}, [], []
    for qi, q in enumerate(queries_for(item)):
        used_q.append(q)
        to_open = []
        for label, slug in SEARCH_TABS.items():
            url = search_url(slug, q)
            try:
                html = fetch(url)
            except SiteBlocked:
                raise
            except requests.HTTPError as e:
                notes.append(f"{label} «{q}»: HTTP {e.response.status_code}")
                continue
            except Exception as e:                       # обрыв связи и т.п. после всех повторов
                notes.append(f"{label} «{q}»: не загрузилось ({type(e).__name__})")
                print(f"  ⚠ {label}: страница поиска не загрузилась ({type(e).__name__}), иду дальше")
                continue
            items, empty = search_results(html, slug)
            if not items and not empty:
                DEBUG_DIR.mkdir(exist_ok=True)
                dump = DEBUG_DIR / f"search_{item['n']:02d}_{qi + 1}_{slug}.html"
                dump.write_text(html, encoding="utf-8")
                notes.append(f"{label} «{q}»: результаты не распознаны (HTML в {dump})")
            for pos, it in enumerate(items, 1):
                t, a_ok = score_search_item(item, it)
                row = {"№ в перечне": item["n"], "ASTA": item["asta"], "Автор (перечень)": item["author"],
                       "Название (перечень)": item["title"], "Запрос": q, "Вкладка": label, "Позиция": pos,
                       "Название (поиск)": it["name"], "Авторы (поиск)": it["authors"], "Описание": it["desc"],
                       "Ссылка": it["url"], "Похожесть названия (поиск)": t,
                       "Автор в результате": "да" if a_ok else "нет", "Карточка открыта": "нет", "Итог": ""}
                cand_rows.append(row)
                if worth_card(t, a_ok) and it["url"] not in found:
                    to_open.append((t, a_ok, it["url"], f"поиск «{q}»: {label}, поз. {pos}", row))
        # открываем карточки самых похожих кандидатов
        to_open.sort(key=lambda x: (-(x[0] + 0.3 * x[1])))
        opened = 0
        for t, a_ok, url, how, row in to_open:
            if url in found:
                found[url]["how"].append(how)
                row["Карточка открыта"], row["Итог"] = "да", found[url]["level"]
                continue
            if opened >= MAX_CARDS_PER_BOOK:
                row["Итог"] = "не открыта (лимит)"
                continue
            opened += 1
            b = get_card(url, cards)
            ct, ca, lvl = judge_card(item, b) if not b.get("error") else (t, a_ok, "проверить")
            row["Карточка открыта"], row["Итог"] = "да", lvl
            if lvl != "не подходит":
                found[url] = {"url": url, "how": [how], "t": ct, "a_ok": ca, "level": lvl}
        if any(f["level"] in ("точно", "вероятно") for f in found.values()):
            break                                            # следующий вариант запроса не нужен

    # парные версии и «Другие версии» с карточек уверенных совпадений
    if want_extra:
        extra = 0
        queue = [u for u, f in found.items() if f["level"] in ("точно", "вероятно")]
        while queue and extra < MAX_EXTRA_PER_BOOK:
            src = queue.pop(0)
            b = cards.get(src) or {}
            links = ([b["pair_url"]] if b.get("pair_url") else []) + list(b.get("versions_urls") or [])
            for u in links:
                if u in found or extra >= MAX_EXTRA_PER_BOOK:
                    continue
                extra += 1
                nb = get_card(u, cards)
                if nb.get("error"):
                    continue
                ct, ca, lvl = judge_card(item, nb)
                kind = "парная версия" if u == b.get("pair_url") else "«Другие версии»"
                if lvl != "не подходит":
                    found[u] = {"url": u, "how": [f"с карточки {b.get('edition_id')}: {kind}"],
                                "t": ct, "a_ok": ca, "level": lvl}
                    if lvl in ("точно", "вероятно"):
                        queue.append(u)
                else:
                    notes.append(f"{kind} {nb.get('edition_id')} «{nb.get('title')}» — не подходит, не взята")
    return list(found.values()), used_q, notes


def main():
    started = datetime.now()
    ask_clear_cache()
    items, source = load_list()
    print(f"Книг в перечне: {len(items)}")
    want_extra = ask_yes("Добирать с найденных карточек парные версии (текст ↔ аудио) и «Другие версии книги»?")

    cards, cand_rows, book_rows, list_rows = {}, [], [], []
    stopped = None
    for item in items:
        print(f"\n[{item['n']}/{len(items)}] {item['author']} — «{item['title']}» ({item['asta']})")
        row = {"№": item["n"], "ASTA": item["asta"], "Автор": item["author"], "Название": item["title"],
               "Статус поиска": "", "Изданий найдено": 0, "текст": 0, "аудио": 0, "комиксы": 0,
               "Наших изданий": 0, "Только «проверить»": 0, "Запросы": "", "Ссылки найденных": "", "Примечания": ""}
        list_rows.append(row)
        if stopped:
            row["Статус поиска"] = f"не обработано ({stopped})"
            continue
        try:
            found, used_q, notes = process_item(item, cards, cand_rows, want_extra)
        except (KeyboardInterrupt, SiteBlocked) as e:
            stopped = "прервано" if isinstance(e, KeyboardInterrupt) else "сайт ограничил запросы"
            row["Статус поиска"] = f"не обработано ({stopped})"
            print(f"\n⛔ {stopped.capitalize()}. Остальные книги пропущены, собранное сохраняется.")
            continue
        found.sort(key=lambda f: (LEVEL_ORDER[f["level"]],
                                  {"текст": 0, "аудио": 1, "комикс": 2}.get(cards[f["url"]].get("type"), 3)))
        for f in found:
            b = cards[f["url"]]
            book_rows.append({**b, "n": item["n"], "asta": item["asta"], "list_author": item["author"],
                              "list_title": item["title"], "level": f["level"], "title_score": f["t"],
                              "author_ok": f["a_ok"], "how": "; ".join(f["how"]),
                              "group": publisher_group(b.get("rightholder"), b.get("publisher"))})
        sure = [f for f in found if f["level"] in ("точно", "вероятно")]
        types = [cards[f["url"]].get("type") for f in sure]
        row.update({
            "Статус поиска": ("найдено" if sure else ("только «проверить»" if found else "не найдено")),
            "Изданий найдено": len(sure), "текст": types.count("текст"), "аудио": types.count("аудио"),
            "комиксы": types.count("комикс"),
            "Наших изданий": sum(publisher_group(cards[f["url"]].get("rightholder"),
                                                 cards[f["url"]].get("publisher")) == "Наши" for f in sure),
            "Только «проверить»": len(found) - len(sure), "Запросы": " | ".join(used_q),
            "Ссылки найденных": "; ".join(f["url"] for f in sure), "Примечания": "; ".join(notes)})
        print(f"  {row['Статус поиска']}: изданий {len(sure)} (текст {row['текст']}, аудио {row['аудио']}, "
              f"комиксы {row['комиксы']}; наших {row['Наших изданий']})"
              + (f", ещё «проверить»: {row['Только «проверить»']}" if row["Только «проверить»"] else "")
              + f"  ({net_summary()})")

    about = [
        ("Программа", f"{PROGRAM}, версия {VERSION}"),
        ("Проект", PROJECT_NAME), ("Чат", CHAT_NAME),
        ("Запуск", started.strftime("%Y-%m-%d %H:%M:%S")),
        ("Перечень", f"{source}, книг: {len(items)}"),
        ("Вкладки поиска", ", ".join(SEARCH_TABS)),
        ("Результатов с вкладки", RESULTS_PER_TAB),
        ("Парные версии и «Другие версии»", "добирались" if want_extra else "не добирались"),
        ("Пороги названия (точно / вероятно / частично)", f"{TITLE_EXACT} / {TITLE_LIKELY} / {TITLE_PARTIAL}"),
        ("Режим", "Инкогнито (без cookies)" if INCOGNITO else "с cookies"),
        ("Запросы", net_summary()),
        ("Найдено книг перечня", sum(r["Статус поиска"] == "найдено" for r in list_rows)),
        ("Изданий в листе «Книги»", len(book_rows)),
        ("Остановка", stopped or "нет"),
    ]
    path, _ = to_excel(list_rows, book_rows, cand_rows, about)
    ok = sum(r["Статус поиска"] == "найдено" for r in list_rows)
    print(f"\nГотово: найдено {ok} из {len(items)} книг перечня, изданий в листе «Книги»: {len(book_rows)}.")
    print(f"Запросы: {net_summary()}")
    if STATS["ограничений"]:
        print("⚠ Сайт ограничивал запросы. Перед следующим запуском лучше сделать перерыв на 1–2 часа.")
    print(f"Файл: {path.resolve()}")
    return pd.DataFrame(list_rows)


if __name__ == "__main__":
    result = main()
    try:
        print(result)   # в Jupyter покажет таблицу
    except NameError:
        pass
    
