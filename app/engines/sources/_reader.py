"""
_reader.py — общий разбор листа-источника в строки-словари.

Переиспользует проверенное определение колонок по алиасам заголовков из
fill_optiauto.py (одна реализация на все адаптеры — без дублирования):
  • _detect_source_columns — находит строку заголовка и колонки по алиасам;
  • _read_source_rows      — читает строки данных, отсеивая итоги/пустые.

Здесь — только мост «строки-словари → canonical.Good/Box» и определение валюты.
"""
from __future__ import annotations

# Проверенные детекторы и словари алиасов живут в fill_optiauto (не дублируем).
from ..fill_optiauto import (          # noqa: E402
    _detect_source_columns,
    _read_source_rows,
    INV_SOURCE_ALIASES,
    PL_SOURCE_ALIASES,
)
from ..canonical import Good, Box


def read_goods(ws) -> list[Good]:
    """Лист инвойса → список Good."""
    cols, hdr = _detect_source_columns(ws, INV_SOURCE_ALIASES)
    goods: list[Good] = []
    for rec in _read_source_rows(ws, cols, hdr):
        goods.append(Good(
            brand=rec.get("brand"),
            part_number=rec.get("part_number"),
            description_en=rec.get("description_en"),
            description_ru=rec.get("description"),   # алиас 'description' = ТОЛЬКО русское
            qty=rec.get("qty"),
            unit_weight=rec.get("unit_weight"),
            unit_price=rec.get("unit_s_price"),
            amount=rec.get("amt_aed"),
            coo=rec.get("coo"),
            hs_code=rec.get("hs_code"),
            incoming_decl=rec.get("incoming_decl"),
        ))
    return goods


def read_boxes(ws) -> list[Box]:
    """Лист упаковки → список Box (сканируем всю книгу: PL-заголовок бывает ниже)."""
    cols, hdr = _detect_source_columns(ws, PL_SOURCE_ALIASES, scan_rows=None)
    boxes: list[Box] = []
    for rec in _read_source_rows(ws, cols, hdr):
        boxes.append(Box(
            number=rec.get("pkg_number"),
            width=rec.get("width"),
            height=rec.get("height"),
            length=rec.get("length"),
            qty=rec.get("qty"),
            gross_weight=rec.get("gross_weight"),
            cbm=rec.get("cbm"),
            actual_weight=rec.get("actual_weight"),
        ))
    return boxes


def sniff_currency(ws, scan_rows: int = 30) -> str | None:
    """Определить валюту по заголовкам инвойса (AED / USD)."""
    limit = min(scan_rows, ws.max_row or 1)
    for r in range(1, limit + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            lbl = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if "aed" in lbl:
                return "AED"
            if "usd" in lbl:
                return "USD"
    return None
