#!/usr/bin/env python3
"""Завантажує YML-фід Прому і будує файл у форматі експорту Прому.

    PROM_FEED_URL=... python3 fetch_feed.py prom_feed.xlsx

Код запчастини, виробник і сумісність беремо з <param> і <vendor> фіду,
а чого там немає - з prom_base.json (повний експорт Прому, за ID товару).
"""
import json
import os
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET

import openpyxl

from fetch_prom import CHAR_NAMES, MIN_SHARE

QTY_TAGS = ("quantity_in_stock", "quantityInStock", "stock_quantity", "quantity")


def download(url):
    for attempt in range(5):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "avtopro-feed"})
            with urllib.request.urlopen(req, timeout=180) as resp:
                return resp.read()
        except Exception as e:
            if attempt == 4:
                raise
            print("повтор:", e)
            time.sleep(10 * (attempt + 1))


def text(el, *tags):
    for t in tags:
        v = el.findtext(t)
        if v and v.strip():
            return v.strip()
    return None


def main(dst):
    url = os.environ.get("PROM_FEED_URL")
    if not url:
        sys.exit("Немає секрету PROM_FEED_URL")
    root = ET.fromstring(download(url))
    shop = root.find("shop") if root.tag != "shop" else root
    groups = {c.get("id"): (c.text or "").strip() for c in shop.iter("category")}
    offers = list(shop.iter("offer"))
    base = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "prom_base.json")))
    print(f"Фід: товарів {len(offers)}, груп {len(groups)}")
    if offers:
        print("теги товару у фіді:", sorted({c.tag for c in offers[0]}), "атрибути:", sorted(offers[0].attrib))
    if len(offers) < MIN_SHARE * len(base):
        sys.exit(f"У фіді лише {len(offers)} товарів (у базі {len(base)}) - файли не оновлюю")

    header = ["Код_товару", "Назва_позиції", "Назва_позиції_укр", "Опис", "Опис_укр", "Ціна",
              "Посилання_зображення", "Наявність", "Кількість", "Назва_групи",
              "Унікальний_ідентифікатор", "Виробник"]
    for _ in CHAR_NAMES:
        header += ["Назва_Характеристики", "Одиниця_виміру_Характеристики", "Значення_Характеристики"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export Products Sheet"
    ws.append(header)
    no_group = 0
    for o in offers:
        pid = o.get("id")
        known = base.get(str(pid), {})
        params = {}
        for p in o.findall("param"):
            n = (p.get("name") or "").strip()
            if n and n not in params and p.text and p.text.strip():
                params[n] = p.text.strip()
        avail = (o.get("available") or "").lower()
        qty = text(o, *QTY_TAGS)
        try:
            qty = int(float(qty)) if qty is not None else None
        except ValueError:
            qty = None
        in_stock = avail == "true" and qty != 0
        group = groups.get(text(o, "categoryId") or "") or known.get("_group")
        if not group:
            no_group += 1
        pics = [p.text.strip() for p in o.findall("picture") if p.text and p.text.strip()]
        name = text(o, "name_ua", "name")
        row = [text(o, "vendorCode"), name, name, text(o, "description_ua", "description"), None,
               text(o, "price"), ", ".join(pics), "+" if in_stock else "-", qty, group, pid,
               known.get("_brand") or text(o, "vendor")]
        for n in CHAR_NAMES:
            val = params.get(n) or known.get(n)
            if n == "Сумісність з маркою" and not val:
                val = params.get("Совместимость с маркой") or known.get("Совместимость с маркой")
            row += [n if val else None, None, val]
        ws.append(row)
    if no_group:
        print(f"увага: {no_group} товарів без групи")
    wb.save(dst)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "prom_feed.xlsx")
