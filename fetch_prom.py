#!/usr/bin/env python3
"""Завантажує товари з Prom.ua API і будує файл у форматі експорту Прому.

    PROM_TOKEN=... python3 fetch_prom.py prom_api.xlsx

Код запчастини, виробник і сумісність беремо з prom_base.json (з повного
експорту Прому, за ID товару), бо API їх не віддає. Ціну, наявність, фото
і групу - свіжі з API. Для нових товарів код = артикул (sku) з Прому.
"""
import json
import os
import sys
import time
import urllib.request

import openpyxl

API = "https://my.prom.ua/api/v1/products/list"
MIN_SHARE = 0.5  # якщо прийшло менше половини товарів - щось не так, файли не оновлюємо
CHAR_NAMES = ["Код запчастини", "Виробник", "Сумісність з маркою", "Сумісність з моделлю", "Стан"]


def fetch_all(token):
    products, last_id = {}, None
    while True:
        url = API + "?limit=100" + (f"&last_id={last_id}" if last_id else "")
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}",
                                                   "Accept": "application/json"})
        for attempt in range(5):
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    batch = json.load(resp).get("products", [])
                break
            except Exception as e:  # мережа/ліміт - чекаємо і пробуємо ще
                if attempt == 4:
                    raise
                print("повтор:", e)
                time.sleep(5 * (attempt + 1))
        new = [p for p in batch if p["id"] not in products]
        if not new:
            return list(products.values())
        if not products:
            print("поля товару в API:", sorted(new[0].keys()))
        for p in new:
            products[p["id"]] = p
        last_id = min(p["id"] for p in new) - 1
        time.sleep(0.5)


def presence_mark(p):
    if p.get("presence") in ("available", "order") or p.get("in_stock") is True:
        return "+"
    return "-"


def main(dst):
    token = os.environ.get("PROM_TOKEN")
    if not token:
        sys.exit("Немає секрету PROM_TOKEN")
    base = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "prom_base.json")))
    products = fetch_all(token)
    shown = [p for p in products if p.get("status") in (None, "on_display")]
    print(f"API: товарів {len(products)}, на сайті {len(shown)}")
    if len(shown) < MIN_SHARE * len(base):
        sys.exit(f"Отримано лише {len(shown)} товарів (у базі {len(base)}) - файли не оновлюю")

    header = ["Код_товару", "Назва_позиції", "Назва_позиції_укр", "Опис", "Опис_укр", "Ціна",
              "Посилання_зображення", "Наявність", "Кількість", "Назва_групи",
              "Унікальний_ідентифікатор", "Виробник"]
    for _ in CHAR_NAMES:
        header += ["Назва_Характеристики", "Одиниця_виміру_Характеристики", "Значення_Характеристики"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export Products Sheet"
    ws.append(header)
    for p in shown:
        known = base.get(str(p["id"]), {})
        names = p.get("name_multilang") or {}
        descs = p.get("description_multilang") or {}
        images = [i.get("url") for i in p.get("images") or [] if i.get("url")]
        if not images and p.get("main_image"):
            images = [p["main_image"]]
        qty = p.get("quantity_in_stock")
        row = [p.get("sku"), p.get("name"), names.get("uk"), p.get("description"), descs.get("uk"),
               p.get("price"), ", ".join(images), presence_mark(p), qty if qty else None,
               (p.get("group") or {}).get("name"), p["id"], known.get("_brand")]
        for n in CHAR_NAMES:
            val = known.get(n)
            if n == "Сумісність з маркою" and not val:
                val = known.get("Совместимость с маркой")
            row += [n if val else None, None, val]
        ws.append(row)
    wb.save(dst)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "prom_api.xlsx")
