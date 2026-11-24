#!/usr/bin/env python3
"""Повний експорт Prom.ua -> окремий файл Автопро на кожен склад (групу Прому).

Використання:
    python3 split_by_group.py prom_export.xlsx output_dir/

Результат: output_dir/<склад>.xlsx (для Автопро) і output_dir/propushcheni/<склад>.xlsx.
"""
import os
import re
import sys
import tempfile
from collections import defaultdict

import openpyxl

import prom_to_avtopro as conv

# Назва групи в Промі -> назва файлу (= склад на Автопро). Шукаємо за ключовими словами.
WAREHOUSES = [
    (("megane",), "megane3"),
    (("stepway",), "sandero_stepway"),
    (("sandero 2", "рестайл"), "sandero2_restyling"),
    (("sandero 2",), "sandero2"),
    (("sandero",), "sandero1"),
    (("accent",), "accent4"),
    (("logan mcv",), "logan_mcv"),
    (("logan 2",), "logan2"),
    (("logan", "рестайл"), "logan1_restyling"),
    (("logan",), "logan1"),
    (("dokker",), "dokker"),
    (("symbol",), "symbol2"),
    (("duster",), "duster"),
    (("fabia", "універсал"), "fabia2_universal"),
    (("fabia",), "fabia2_hatchback"),
]


def warehouse_for(group):
    g = (group or "").lower()
    for keys, name in WAREHOUSES:
        if all(k in g for k in keys):
            return name
    slug = re.sub(r"[^a-z0-9]+", "_", g.encode("ascii", "ignore").decode()).strip("_")
    return slug or "bez_grupy"


def main(src, out_dir):
    os.makedirs(os.path.join(out_dir, "propushcheni"), exist_ok=True)
    ws = openpyxl.load_workbook(src, read_only=True)["Export Products Sheet"]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    gi = header.index("Назва_групи")
    groups = defaultdict(list)
    names = {}
    for r in rows:
        wh = warehouse_for(r[gi])
        groups[wh].append(r)
        names.setdefault(wh, r[gi])

    with tempfile.TemporaryDirectory() as tmp:
        for wh, items in sorted(groups.items()):
            part = os.path.join(tmp, f"{wh}.xlsx")
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = "Export Products Sheet"
            sheet.append(header)
            for r in items:
                sheet.append(r)
            wb.save(part)
            dst = os.path.join(out_dir, f"{wh}.xlsx")
            out, skipped, merged = conv.convert(part, dst)
            os.replace(dst.rsplit(".", 1)[0] + "_propushcheni.xlsx",
                       os.path.join(out_dir, "propushcheni", f"{wh}.xlsx"))
            print(f"{wh}\t{len(out)}\tпропущено {len(skipped)}\t{names[wh]}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
