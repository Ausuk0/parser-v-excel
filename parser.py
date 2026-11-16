"""
Парсер каталога books.toscrape.com с выгрузкой в Excel.

Собирает по страницам каталога: название, категорию, цену, наличие
(и остаток в штуках), рейтинг (1-5) и ссылку на товар. Результат
сохраняется в .xlsx с двумя листами: «Товары» и «Сводка».

Примеры:
    python parser.py
    python parser.py --pages 5
    python parser.py --category "Science Fiction" --out scifi.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/"
CATALOG_URL = urljoin(BASE_URL, "catalogue/page-1.html")
USER_AGENT = "Mozilla/5.0 (compatible; catalog-parser-demo/1.0)"

RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# Консоль Windows иногда не умеет печатать часть символов — не падаем из-за этого.
for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


@dataclass
class Product:
    title: str
    category: str
    price: float
    in_stock: bool
    quantity: int
    rating: int
    url: str


class Fetcher:
    """HTTP-клиент с вежливой задержкой и повторными попытками."""

    def __init__(self, delay: float = 0.5, retries: int = 3, timeout: float = 20):
        self.delay = delay
        self.retries = retries
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT
        self._last_request = 0.0

    def get_soup(self, url: str) -> BeautifulSoup:
        last_error: Exception | None = None
        for attempt in range(1, self.retries + 1):
            wait = self.delay - (time.monotonic() - self._last_request)
            if wait > 0:
                time.sleep(wait)
            self._last_request = time.monotonic()
            try:
                resp = self.session.get(url, timeout=self.timeout)
                if resp.status_code == 404:
                    raise LookupError(f"Страница не найдена (404): {url}")
                resp.raise_for_status()
                # Сайт отдаёт UTF-8, но не указывает кодировку в заголовке.
                return BeautifulSoup(resp.content, "html.parser", from_encoding="utf-8")
            except LookupError:
                raise
            except requests.RequestException as exc:
                last_error = exc
                if attempt < self.retries:
                    pause = 2 ** attempt
                    print(f"    ! Ошибка сети ({exc.__class__.__name__}), "
                          f"повтор {attempt}/{self.retries - 1} через {pause} с")
                    time.sleep(pause)
        raise ConnectionError(f"Не удалось загрузить {url}: {last_error}")


def get_categories(fetcher: Fetcher) -> dict[str, str]:
    """Возвращает {название категории: ссылка на первую страницу}."""
    soup = fetcher.get_soup(BASE_URL + "index.html")
    links = soup.select("div.side_categories ul li ul li a")
    return {a.get_text(strip=True): urljoin(BASE_URL, a["href"]) for a in links}


def parse_listing(soup: BeautifulSoup, page_url: str) -> tuple[list[str], str | None, int | None]:
    """Ссылки на товары со страницы, ссылка на следующую страницу и общее число страниц."""
    links = [urljoin(page_url, a["href"]) for a in soup.select("article.product_pod h3 a")]
    next_a = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_a["href"]) if next_a else None
    total = None
    current = soup.select_one("li.current")
    if current:
        m = re.search(r"of\s+(\d+)", current.get_text())
        if m:
            total = int(m.group(1))
    elif next_url is None:
        total = 1
    return links, next_url, total


def parse_product(soup: BeautifulSoup, url: str) -> Product:
    main = soup.select_one("div.product_main")
    title = main.h1.get_text(strip=True)

    price_text = main.select_one("p.price_color").get_text(strip=True)
    price = float(re.sub(r"[^\d.]", "", price_text))

    avail_text = main.select_one("p.availability").get_text(" ", strip=True)
    m = re.search(r"\((\d+)\s+available\)", avail_text)
    quantity = int(m.group(1)) if m else 0
    in_stock = "in stock" in avail_text.lower() and (quantity > 0 or m is None)

    rating_tag = main.select_one("p.star-rating")
    rating = 0
    if rating_tag:
        for cls in rating_tag.get("class", []):
            rating = RATING_WORDS.get(cls, rating)

    crumbs = soup.select("ul.breadcrumb li a")
    category = crumbs[-1].get_text(strip=True) if len(crumbs) >= 3 else "Без категории"

    return Product(title, category, price, in_stock, quantity, rating, url)


def scrape(fetcher: Fetcher, start_url: str, max_pages: int | None) -> list[Product]:
    products: list[Product] = []
    page_url: str | None = start_url
    page_no = 0
    total_pages: int | None = None
    started = time.monotonic()

    try:
        while page_url and (max_pages is None or page_no < max_pages):
            page_no += 1
            soup = fetcher.get_soup(page_url)
            links, next_url, total = parse_listing(soup, page_url)
            total_pages = total_pages or total
            known = [x for x in (total_pages, max_pages) if x]
            planned = min(known) if known else "?"
            print(f"\nСтраница {page_no}/{planned}: найдено товаров {len(links)}")

            for i, link in enumerate(links, 1):
                try:
                    product = parse_product(fetcher.get_soup(link), link)
                except (ConnectionError, LookupError, AttributeError, ValueError) as exc:
                    print(f"  [{i:>2}/{len(links)}] пропущен: {exc}")
                    continue
                products.append(product)
                short = product.title if len(product.title) <= 55 else product.title[:52] + "..."
                print(f"  [{i:>2}/{len(links)}] {short}  |  £{product.price:.2f}  |  {product.category}")

            elapsed = time.monotonic() - started
            print(f"  Итого собрано: {len(products)} товаров, прошло {elapsed:.0f} с")
            page_url = next_url
    except KeyboardInterrupt:
        print("\nОстановлено пользователем — сохраняю то, что успели собрать.")

    return products


# ---------------------------------------------------------------- Excel

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PRICE_FORMAT = '#,##0.00 "£"'


def style_header(ws, columns: int) -> None:
    for col in range(1, columns + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def write_excel(products: list[Product], out_path: Path, source: str) -> None:
    wb = Workbook()

    # --- Лист «Товары»
    ws = wb.active
    ws.title = "Товары"
    headers = ["№", "Название", "Категория", "Цена", "Наличие", "Остаток, шт.", "Рейтинг (1–5)", "Ссылка"]
    widths = [6, 60, 22, 12, 15, 13, 14, 50]
    ws.append(headers)
    style_header(ws, len(headers))

    for n, p in enumerate(products, 1):
        ws.append([n, p.title, p.category, p.price,
                   "В наличии" if p.in_stock else "Нет в наличии",
                   p.quantity, p.rating, p.url])
        row = ws.max_row
        ws.cell(row, 4).number_format = PRICE_FORMAT
        link = ws.cell(row, 8)
        link.hyperlink = p.url
        link.style = "Hyperlink"
        for col in (1, 4, 5, 6, 7):
            ws.cell(row, col).alignment = Alignment(horizontal="center")
        ws.cell(row, 4).alignment = Alignment(horizontal="right")

    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"

    # --- Лист «Сводка»
    ws2 = wb.create_sheet("Сводка")
    ws2.append(["Категория", "Кол-во товаров", "Средняя цена", "Мин. цена", "Макс. цена"])
    style_header(ws2, 5)

    groups: dict[str, list[float]] = defaultdict(list)
    for p in products:
        groups[p.category].append(p.price)
    for cat in sorted(groups, key=lambda c: (-len(groups[c]), c)):
        prices = groups[cat]
        ws2.append([cat, len(prices), round(sum(prices) / len(prices), 2), min(prices), max(prices)])

    all_prices = [p.price for p in products]
    ws2.append(["ИТОГО", len(all_prices),
                round(sum(all_prices) / len(all_prices), 2) if all_prices else 0,
                min(all_prices, default=0), max(all_prices, default=0)])
    total_row = ws2.max_row
    for col in range(1, 6):
        c = ws2.cell(total_row, col)
        c.font = Font(bold=True)
        c.fill = TOTAL_FILL
    for row in ws2.iter_rows(min_row=2, max_row=total_row):
        row[1].alignment = Alignment(horizontal="center")
        for c in row[2:]:
            c.number_format = PRICE_FORMAT
        for c in row:
            c.border = BORDER
    for col, w in zip("ABCDE", (28, 16, 15, 13, 13)):
        ws2.column_dimensions[col].width = w

    info_row = total_row + 2
    ws2.cell(info_row, 1, f"Источник: {source}").font = Font(italic=True, color="808080")
    ws2.cell(info_row + 1, 1, f"Дата выгрузки: {datetime.now():%d.%m.%Y %H:%M}").font = Font(italic=True, color="808080")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------------------------------------------------------------- CLI

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Парсер каталога books.toscrape.com с выгрузкой в Excel.")
    ap.add_argument("--pages", type=int, default=None, metavar="N",
                    help="сколько страниц каталога обработать (по умолчанию — все)")
    ap.add_argument("--category", default=None, metavar="ИМЯ",
                    help='только одна категория, например "Travel" или "Science Fiction"')
    ap.add_argument("--out", default=None, metavar="ФАЙЛ",
                    help="путь к итоговому .xlsx (по умолчанию books_ДАТА_ВРЕМЯ.xlsx)")
    ap.add_argument("--delay", type=float, default=0.5, metavar="СЕК",
                    help="пауза между запросами в секундах (по умолчанию 0.5)")
    args = ap.parse_args(argv)
    if args.pages is not None and args.pages < 1:
        ap.error("--pages должно быть не меньше 1")
    if args.delay < 0:
        ap.error("--delay не может быть отрицательным")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    out_path = Path(args.out) if args.out else Path(f"books_{datetime.now():%Y-%m-%d_%H-%M}.xlsx")
    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")

    fetcher = Fetcher(delay=args.delay)

    try:
        if args.category:
            categories = get_categories(fetcher)
            lookup = {name.lower(): (name, url) for name, url in categories.items()}
            found = lookup.get(args.category.strip().lower())
            if not found:
                print(f"Категория «{args.category}» не найдена. Доступные категории:")
                print(", ".join(sorted(categories)))
                return 2
            name, start_url = found
            print(f"Категория: {name}")
        else:
            start_url = CATALOG_URL
            print("Весь каталог")
        print(f"Лимит страниц: {args.pages or 'нет'}; пауза между запросами: {args.delay} с")

        products = scrape(fetcher, start_url, args.pages)
    except (ConnectionError, LookupError) as exc:
        print(f"\nОшибка: {exc}")
        return 1

    if not products:
        print("\nНичего не собрано — файл не создан.")
        return 1

    write_excel(products, out_path, start_url)
    print(f"\nГотово: {len(products)} товаров сохранено в {out_path.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
