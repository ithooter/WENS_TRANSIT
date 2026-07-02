"""
combined.py — адаптер «Combined» (1 файл-исходник).

Один .xlsx, где инвойс и упаковка лежат на разных листах одной книги
(классически Sheet1 = товары, Sheet2 = коробки, но имена могут быть любыми —
лист ищется нечётко по имени + сигнатуре колонок через sheet_resolver).
"""
from __future__ import annotations

from pathlib import Path

import openpyxl

from ..canonical import Requisites, Shipment
from ..sheet_resolver import resolve_sheet
from ._reader import read_boxes, read_goods, sniff_currency


class CombinedAdapter:
    name = "combined"

    def matches(self, paths: list[Path]) -> bool:
        # Один файл-книга. (Многофайловые форматы — отдельные адаптеры, Фаза 2.)
        return len(paths) == 1 and paths[0].suffix.lower() == ".xlsx"

    def read(self, paths: list[Path]) -> Shipment:
        wb = openpyxl.load_workbook(paths[0], data_only=True)
        try:
            inv_name = resolve_sheet(wb, "invoice")
            pl_name = resolve_sheet(wb, "packing", exclude={inv_name})
            inv_ws = wb[inv_name]
            goods = read_goods(inv_ws)
            boxes = read_boxes(wb[pl_name]) if pl_name else []
            currency = sniff_currency(inv_ws)
        finally:
            wb.close()

        return Shipment(
            goods=goods,
            boxes=boxes,
            requisites=Requisites(currency=currency),
            source_format=self.name,
        )
