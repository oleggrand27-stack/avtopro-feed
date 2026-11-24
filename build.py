#!/usr/bin/env python3
"""API Прому -> 15 файлів для Автопро у папці site/ + сторінка з посиланнями."""
import datetime
import os
import sys

import fetch_feed
import fetch_prom
import split_by_group

# Основне джерело - YML-фід Прому; API-токен - запасний варіант
if os.environ.get("PROM_FEED_URL"):
    fetch_feed.main("prom_api.xlsx")
else:
    fetch_prom.main("prom_api.xlsx")
split_by_group.main("prom_api.xlsx", "site")
files = sorted(f for f in os.listdir("site") if f.endswith(".xlsx"))
if not files:
    sys.exit("Порожній результат")
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=2))).strftime("%d.%m.%Y %H:%M")
links = "\n".join(f'<li><a href="{f}">{f}</a></li>' for f in files)
with open("site/index.html", "w") as fh:
    fh.write(f'<!doctype html><meta charset="utf-8"><title>Автопро склади</title>'
             f'<h1>Файли для Автопро</h1><p>Оновлено: {now}</p><ul>{links}</ul>')
