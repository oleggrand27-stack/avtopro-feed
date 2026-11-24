#!/usr/bin/env python3
"""Конвертер експорту товарів Prom.ua у прайс для імпорту на Автопро (avtopro.ua).

Використання (результат: прайс без заголовків + файл *_propushcheni.xlsx):
    python3 prom_to_avtopro.py prom_export.xlsx avtopro_import.xlsx

Потрібна бібліотека openpyxl (pip install openpyxl).
"""
import html
import re
import sys
from collections import Counter

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Колонки вихідного файлу за вимогами Автопро (рядка заголовків у файлі немає):
# A Виробник, B Код, C Опис, D Ціна, E Кількість, F Б/В, G Фото
OUT_COLUMNS = ["Виробник", "Код", "Опис", "Ціна", "Кількість", "Б/В", "Фото"]

IN_STOCK = {"+", "!"}          # "+" і "!" у Промі = товар є
BAD_BRANDS = {"logan", "mcv", "a+", ""}  # це не виробники, а помилки в полі "Виробник"
DESC_LIMIT = 1000
# Оригінальні номери Dacia - це номери Renault; в каталозі Автопро вони під Renault.
# Без цього та сама деталь як Dacia і як Renault дає «Дублікати» і «Не знайдено».
BRAND_ALIASES = {"dacia": "Renault"}
# Оригінальні номери Renault (8200xxxxxx, 7700xxxxxx, 123456789R) - завжди Renault,
# навіть якщо в Промі помилково вказано інший виробник (ABS, Skoda тощо).
CAR_MAKES = {"renault", "dacia", "hyundai", "skoda", "vag", "volkswagen", "kia"}
RENAULT_CODE = re.compile(r"^(\d{9}R|82\d{8}|77\d{8})$")

# Кириличні літери, схожі на латинські, які трапляються в кодах (напр. 8200792584Р)
CYR2LAT = str.maketrans("АВСЕНІКМОРТХаеорсхі", "ABCEHIKMOPTXaeopcxi")


def clean_text(value):
    if not value:
        return ""
    text = re.sub(r"<br\s*/?>|</p>|</li>", "\n", str(value), flags=re.I)
    text = html.unescape(re.sub(r"<[^>]+>", "", text)).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n\s*\n+", "\n", text).strip()


PROMO_PATTERNS = [
    r'"?\s*(?:\b[\w-]+="[^"]*"\s*)+>?',              # залишки HTML-атрибутів (data-ved="..." ...>)
    r"https?://\S+|www\.\S+",                      # посилання на сайт
    r"(\+?38)?\s?\(?0\d{2}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}(?!\d)",  # телефони
    r"Наш (телефон|сайт)\s*:?",
    r"У наявності запчастини для[^.]*\.?",
    r"[\U0001F300-\U0001FAFF\u2600-\u27BF\uFE0F]",   # емодзі
]


def strip_promo(text):
    """Прибирає з опису контакти, посилання і рекламу магазину."""
    for pat in PROMO_PATTERNS:
        text = re.sub(pat, " ", text)
    return re.sub(r"\s+", " ", text).strip(" .,;")


def fitment(marks, models):
    """'Renault' + 'Megane|Renault Megane|Scenic' -> 'Renault Megane, Scenic'."""
    brands = [m.strip() for m in str(marks or "").split("|") if m.strip()]
    seen, out = set(), []
    for m in str(models or "").split("|"):
        m = m.strip()
        for b in brands:
            if m.lower().startswith(b.lower() + " "):
                m = m[len(b) + 1:].strip()
        if m and m.lower() not in seen:
            seen.add(m.lower())
            out.append(m)
    head = ", ".join(brands)
    return f"{head} {', '.join(out)}".strip()


MIN_CODE_LEN = 5  # коротші коди (001, 003, "-") - внутрішні номери, не артикули


def code_key(code):
    """Ключ для пошуку дублікатів: Автопро ігнорує пробіли, дефіси, крапки."""
    return re.sub(r"[^0-9A-Z]", "", code.upper())


def clean_code(value):
    """Перший справжній артикул з поля. '769140022R, 6311140012' -> '769140022R'."""
    if value is None:
        return ""
    text = html.unescape(html.unescape(str(value))).replace("\xa0", " ")
    text = text.translate(CYR2LAT).upper().strip()
    for token in re.split(r"[\s,;/]+", text):
        if len(code_key(token)) >= MIN_CODE_LEN:
            return token.strip(".-")
    joined = re.sub(r"\s+", "", text)  # напр. 'TFH P12' -> 'TFHP12'
    return joined if len(code_key(joined)) >= MIN_CODE_LEN else ""


def used_mark(value):
    """Автопро: для вживаних будь-який символ, для нових порожньо."""
    v = (value or "").strip().lower()
    if v in ("новий", "новое", "new"):
        return None
    return 1  # магазин - авторозбірка, за замовчуванням вживане


def read_prom(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["Export Products Sheet"] if "Export Products Sheet" in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    first = {}
    for i, name in enumerate(header):
        first.setdefault(name, i)
    char_start = first.get("Назва_Характеристики", len(header))
    for row in rows:
        get = lambda name: row[first[name]] if name in first and first[name] < len(row) else None
        chars = {}
        for i in range(char_start, len(row) - 2, 3):
            if row[i] and row[i] not in chars:
                chars[str(row[i]).strip()] = row[i + 2]
        yield get, chars


NO_CODE_ONLY = False  # режим --bez-kodu: тільки товари без артикула, з кодом PROM<ID товару>


def convert(src, dst):
    out, skipped = [], []
    for get, chars in read_prom(src):
        name = clean_text(get("Назва_позиції_укр") or get("Назва_позиції"))
        prom_code = get("Код_товару")
        code = clean_code(chars.get("Код запчастини")) or clean_code(prom_code)
        if NO_CODE_ONLY:
            if code:
                continue
            code = f"PROM{get('Ідентифікатор_товару') or get('Унікальний_ідентифікатор')}"
        price = get("Ціна")
        avail = (get("Наявність") or "").strip()
        qty = get("Кількість")

        reason = None
        if avail not in IN_STOCK or qty == 0:
            reason = "немає в наявності"
        elif not price:
            reason = "немає ціни"
        elif not code:
            reason = "немає коректного коду запчастини"
        if reason:
            skipped.append([reason, prom_code, name, price, avail, qty])
            continue

        brand = str(chars.get("Виробник") or get("Виробник") or "").strip()
        marks = str(chars.get("Сумісність з маркою") or chars.get("Совместимость с маркой") or "")
        if brand.lower() in BAD_BRANDS:
            brand = marks.split("|")[0].strip() or "Renault"
        brand = BRAND_ALIASES.get(brand.lower(), brand)
        if RENAULT_CODE.match(code_key(code)):
            brand = "Renault"
        fit = fitment(marks, chars.get("Сумісність з моделлю"))

        try:
            price_val = float(str(price).replace(",", ".").replace(" ", ""))
        except ValueError:
            skipped.append(["некоректна ціна", prom_code, name, price, avail, qty])
            continue

        desc = name
        if fit:
            desc += f". Підходить: {fit}"
        extra = strip_promo(clean_text(get("Опис_укр") or get("Опис")))
        if extra.startswith(name):
            extra = extra[len(name):].strip(" .,;")
        if extra and extra not in desc:
            desc += f". {extra}"
        photos = ";".join(u.strip() for u in str(get("Посилання_зображення") or "").split(",") if u.strip())

        out.append([
            brand,
            code,
            desc[:DESC_LIMIT],
            int(price_val) if price_val.is_integer() else round(price_val, 2),
            max(int(qty or 1), 1),
            used_mark(chars.get("Стан")),
            photos or None,
        ])

    out, merged = merge_duplicates(out)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Прайс"
    for r in out:
        ws.append(r)
    for i, w in enumerate([12, 18, 80, 10, 10, 6, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Коди як текст, щоб Excel не перетворював 0501801 на 501801
    for (cell,) in ws.iter_rows(min_col=2, max_col=2):
        cell.number_format = "@"

    wb.save(dst)

    # Пропущені позиції - окремим файлом, щоб у прайсі не було зайвого
    wb = openpyxl.Workbook()
    sk = wb.active
    sk.title = "Пропущені"
    sk.append(["Причина", "Код товару Prom", "Назва", "Ціна", "Наявність", "Кількість"])
    for r in skipped:
        sk.append(r)
    for c in sk[1]:
        c.font = Font(bold=True)
    wb.save(dst.rsplit(".", 1)[0] + "_propushcheni.xlsx")
    return out, skipped, merged


def merge_duplicates(rows):
    """Автопро не приймає однакові Виробник+Код: склеюємо в один рядок.
    Кількість сумуємо, ціну, опис і фото беремо з найдешевшої позиції."""
    # Бренд складу (найчастіший). Чужі марки авто та сміття типу "AJ", "AE", "ABS"
    # (у Промі це шматки коду чи назви) замінюємо на нього. Номери Renault лишаються Renault.
    if rows:
        main = Counter(r[0] for r in rows).most_common(1)[0][0]
        for r in rows:
            b = r[0]
            junk = len(b) <= 3 and b.upper() not in ("K&N",)
            other_make = b.lower() in CAR_MAKES and b != main
            if (junk or other_make) and not RENAULT_CODE.match(code_key(r[1])):
                r[0] = main
    # Той самий код під різними брендами (напр. Renault і помилково Skoda) -
    # приводимо до найчастішого бренду, щоб не було дублікатів.
    by_code = {}
    for r in rows:
        by_code.setdefault(code_key(r[1]), []).append(r[0])
    for r in rows:
        brands = Counter(by_code[code_key(r[1])])
        if len(brands) > 1:
            top = max(brands.values())
            best = [b for b, n in brands.items() if n == top]
            r[0] = "Renault" if "Renault" in best else best[0]
    groups = {}
    for r in rows:
        groups.setdefault((r[0].lower(), code_key(r[1])), []).append(r)
    out, merged = [], 0
    for items in groups.values():
        best = list(min(items, key=lambda r: r[3]))
        if len(items) > 1:
            merged += len(items) - 1
            best[4] = sum(r[4] for r in items)
            best[6] = best[6] or next((r[6] for r in items if r[6]), None)
        out.append(best)
    return out, merged


if __name__ == "__main__":
    if "--bez-kodu" in sys.argv:
        sys.argv.remove("--bez-kodu")
        NO_CODE_ONLY = True
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    out, skipped, merged = convert(sys.argv[1], sys.argv[2])
    print(f"У прайс: {len(out)}; склеєно дублікатів: {merged}; пропущено: {len(skipped)}")
    for reason, n in Counter(r[0] for r in skipped).most_common():
        print(f"  {reason}: {n}")
