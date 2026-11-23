# -*- coding: utf-8 -*-
"""
Диагностика: что видит воркер на странице ЭК и почему не находится парная аудиоверсия.
Запуск (локально или в GitHub Actions):
  python debug_pair.py https://books.yandex.ru/books/kcHsrwod https://books.yandex.ru/books/GwpW65yB
Для каждой страницы печатает:
  - что распознал parse_book (pair_url, ссылки «Других версий»);
  - все ссылки на издания (подпись, адрес);
  - все элементы с подписью «Аудио» (тег, есть ли href);
  - все uuid аудиокниг, упомянутые где угодно в HTML (в т.ч. во вшитом JSON).
HTML сохраняется в debug_html/<id>.html (в Actions выкладывается как артефакт запуска).
"""
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup

import ybooks_core as yb

OUT = Path("debug_html")


def analyze(html, url):
    soup = BeautifulSoup(html, "lxml")
    card = yb.parse_book(html, url)
    print(f"  размер HTML: {len(html)} симв.; title: {card.get('title')!r}; автор: {card.get('author')!r}")
    print(f"  правообладатель: {card.get('rightholder')!r}")
    print(f"  parse_book -> pair_url: {card.get('pair_url')!r}")
    print(f"  parse_book -> versions_urls: {card.get('versions_urls')!r}")

    print("  Ссылки на издания (подпись -> href):")
    seen = 0
    for a in soup.find_all("a", href=yb.EDITION_RE):
        print(f"    {yb.text_of(a)[:60]!r} -> {a['href']}")
        seen += 1
        if seen >= 25:
            print("    ... (обрезано)")
            break

    print("  Элементы с подписью «Аудио»/«Текст»/«Комикс» (любой тег):")
    found = False
    for el in soup.find_all(True):
        own = "".join(t for t in el.find_all(string=True, recursive=False)).strip()
        if own in ("Аудио", "Текст", "Комикс"):
            found = True
            print(f"    <{el.name}> {own!r} href={el.get('href')!r} родитель=<{el.parent.name if el.parent else None}>")
    if not found:
        print("    (нет ни одного)")

    ids = sorted(set(re.findall(r"audiobooks/([A-Za-z0-9]+)", html)))
    print(f"  uuid аудиокниг, упомянутых в HTML: {ids}")
    print(f"  слово «Аудио» в тексте страницы: {html.count('Аудио')} раз; '/audiobooks/' в HTML: {html.count('/audiobooks/')} раз")


def main():
    urls = sys.argv[1:]
    if not urls:
        sys.exit("Укажите хотя бы один URL карточки.")
    OUT.mkdir(exist_ok=True)
    for url in urls:
        print(f"\n=== {url}")
        try:
            html = yb.fetch(url)
        except yb.SiteBlocked:
            print("  ⛔ Сайт ограничил запросы — останавливаемся.")
            break
        except Exception as e:
            print(f"  ⚠ не открылась: {type(e).__name__}: {e}")
            continue
        (OUT / (url.rstrip("/").rsplit("/", 1)[-1] + ".html")).write_text(html, encoding="utf-8")
        analyze(html, url)


if __name__ == "__main__":
    main()
