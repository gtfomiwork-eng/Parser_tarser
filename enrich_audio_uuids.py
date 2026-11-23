# -*- coding: utf-8 -*-
"""
M1B: подбор uuid для аудиокниг без ссылки — через парную ЭК из БД.

Логика:
  1. Берём аудиокниги (edition_type='audiobooks') с uuid IS NULL, availability='active'.
  2. Для каждой ищем в БД текст/комикс с uuid и ТОЧНО совпадающими
     нормализованными названием (включая числа) и фамилиями авторов.
  3. Кандидат должен быть ровно один, иначе пометка в pair_note.
  4. Открываем карточку ЭК (1 запрос), читаем pair_url.
     Если пара — audiobooks и её uuid свободен, записываем аудиокниге
     uuid + uuid_source='pair'. Заодно обновляем карточку самой ЭК (card).

Запуск (локально или в GitHub Actions):
  SUPABASE_URL=... SUPABASE_SERVICE_KEY=... python enrich_audio_uuids.py --limit 200
--limit — максимум карточек ЭК за один запуск (чекпоинт естественный:
  обработанные аудио либо получают uuid, либо pair_note, и повторно не берутся).
"""
import argparse
import os
import re
import sys
from datetime import datetime, timezone

from supabase import create_client

import ybooks_core as yb

BASE = "https://books.yandex.ru"


def norm_title(s):
    return yb.norm(s)


def norm_authors(s):
    return tuple(sorted(yb.surnames(s or "")))


def numbers(s):
    return set(re.findall(r"\d+", yb.norm(s)))


def fetch_all(build, page=1000):
    """Выгрузка всех строк постранично: Supabase отдаёт не больше 1000 строк за запрос
    (настройка Max Rows), поэтому .limit(5000) / .limit(100000) молча обрезаются."""
    rows, start = [], 0
    while True:
        chunk = build().order("id").range(start, start + page - 1).execute().data
        rows += chunk
        if len(chunk) < page:
            return rows
        start += page


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=200, help="макс. карточек ЭК за запуск")
    args = ap.parse_args()

    url, key = os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_SERVICE_KEY")
    if not url or not key:
        sys.exit("Задайте SUPABASE_URL и SUPABASE_SERVICE_KEY.")
    sb = create_client(url, key)

    # Аудио без uuid и без прежней пометки о неоднозначности
    audios = fetch_all(lambda: sb.table("books").select("id, asta, author, title")
        .eq("edition_type", "audiobooks").is_("uuid", "null")
        .is_("pair_note", "null").eq("availability", "active"))
    print(f"Аудиокниг без uuid: {len(audios)}")

    # Все ЭК/комиксы с uuid — индексируем по (авторы, название)
    texts = fetch_all(lambda: sb.table("books").select("id, asta, author, title, uuid, edition_type")
        .in_("edition_type", ["books", "comicbooks"]).not_.is_("uuid", "null"))
    print(f"ЭК/комиксов с uuid в индексе: {len(texts)}")
    index = {}
    for t in texts:
        index.setdefault((norm_authors(t["author"]), norm_title(t["title"])), []).append(t)

    done = requests_used = 0
    stats = {"linked": 0, "no_pair": 0, "ambiguous": 0, "mismatch": 0, "skipped": 0}
    for a in audios:
        if requests_used >= args.limit:
            print("Достигнут лимит запросов на этот запуск — остальное в следующий раз.")
            break
        key_ = (norm_authors(a["author"]), norm_title(a["title"]))
        if not key_[0] or not key_[1]:
            stats["skipped"] += 1
            continue
        cands = index.get(key_, [])
        # защита по числам в названии (копия правила title_score_n)
        cands = [c for c in cands if numbers(c["title"]) == numbers(a["title"])]
        if not cands:
            stats["skipped"] += 1          # пары в БД нет — найдётся обычным поиском
            continue
        if len(cands) > 1:
            stats["ambiguous"] += 1
            sb.table("books").update({"pair_note": "неоднозначно: несколько ЭК с тем же названием"}) \
                .eq("id", a["id"]).execute()
            continue

        t = cands[0]
        card_url = f"{BASE}/{t['edition_type']}/{t['uuid']}"
        try:
            card = yb.parse_book(yb.fetch(card_url), card_url)
        except yb.SiteBlocked:
            print("⛔ Сайт ограничил запросы — останавливаемся, продолжим в следующий запуск.")
            break
        except Exception as e:
            print(f"  ⚠ {a['asta']}: карточка ЭК не открылась ({type(e).__name__})")
            continue
        requests_used += 1

        # заодно сохраняем свежий снимок карточки ЭК
        sb.table("books").update({
            "card": card, "card_parsed_at": datetime.now(timezone.utc).isoformat(),
        }).eq("id", t["id"]).execute()

        pair_type, pair_uuid = yb.edition_of(card.get("pair_url") or "")
        if pair_type != "audiobooks" or not pair_uuid:
            stats["no_pair"] += 1
            sb.table("books").update({"pair_note": "у ЭК нет парной аудиоверсии на сайте"}) \
                .eq("id", a["id"]).execute()
            continue

        # uuid не должен быть занят другой книгой (один ASTA = один uuid, ТЗ п. 15)
        busy = sb.table("books").select("id").eq("uuid", pair_uuid).neq("id", a["id"]).execute().data
        if busy:
            stats["mismatch"] += 1
            sb.table("books").update({"pair_note": f"uuid пары {pair_uuid} уже занят книгой id={busy[0]['id']}"}) \
                .eq("id", a["id"]).execute()
            continue

        sb.table("books").update({
            "uuid": pair_uuid, "uuid_source": "pair", "pair_note": None,
        }).eq("id", a["id"]).execute()
        stats["linked"] += 1
        done += 1
        print(f"  ✓ {a['asta']} «{a['title']}» -> audiobooks/{pair_uuid}")

    print(f"\nИтог: привязано {stats['linked']}, без пары {stats['no_pair']}, "
          f"неоднозначно {stats['ambiguous']}, конфликт uuid {stats['mismatch']}, "
          f"пропущено {stats['skipped']}. Запросов к сайту: {requests_used} ({yb.net_summary()})")


if __name__ == "__main__":
    main()
