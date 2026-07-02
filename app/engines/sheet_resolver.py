"""
sheet_resolver.py — нечёткий поиск нужного листа в книге.

Проблема: клиент присылает исходник, где лист «Invoice List» назван по-своему
(«Processed Data», «TDSheet», «Sheet1» …). Хардкод-список имён в старом коде
(fill_optiauto._pick_sheet, si.read_source) на такое ломается молча.

Решение: искать лист по РОЛИ по двум признакам —
  1) ключевые слова в ИМЕНИ листа;
  2) СИГНАТУРА КОЛОНОК внутри листа (какие заголовки в нём есть).
Признак по колонкам сильнее имени: даже «Processed Data» опознаётся как инвойс,
если внутри есть Brand+Part Number (или Qty+HS Code).
"""
from __future__ import annotations

# Ключевые слова в имени листа для каждой роли (в нижнем регистре).
_NAME_HINTS = {
    "invoice": ("invoice", "inv", "processed", "tdsheet", "sheet1", "sheet_1",
                "data", "товар", "goods", "items"),
    "packing": ("packing", "pack", "pl", "box", "colli", "sheet2", "sheet_2",
                "упаков", "места", "коробк"),
    "details": ("details", "detail", "req", "реквизит", "sheet2", "info"),
    "proforma": ("proforma", "проформа", "profarma", "pf"),
}

# Токены-заголовки колонок, характерные для роли (частичное вхождение).
_COL_HINTS = {
    "invoice": (("brand", "бренд"), ("part number", "part no", "артикул"),
                ("hscode", "hs code", "тн вэд"), ("quantity", "qty", "кол-во")),
    "packing": (("box n", "box number", "номер коробки", "colli"),
                ("width", "ширина", "dimensions"), ("cbm", "объем", "volume")),
    "proforma": (("артикул",), ("наименование",), ("тнвэд", "тн вэд")),
}


def _norm(v) -> str:
    return str(v).strip().lower() if v is not None else ""


def _column_signature_score(ws, role: str, scan_rows: int = 15) -> int:
    """Сколько характерных для роли групп-токенов встречается в заголовках листа."""
    groups = _COL_HINTS.get(role)
    if not groups:
        return 0
    labels: set[str] = set()
    limit = min(scan_rows, ws.max_row or 1)
    for r in range(1, limit + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            lbl = _norm(ws.cell(row=r, column=c).value)
            if lbl:
                labels.add(lbl)
    score = 0
    for group in groups:
        if any(any(tok in lbl for lbl in labels) for tok in group):
            score += 1
    return score


def _name_score(title: str, role: str) -> int:
    t = _norm(title)
    return sum(1 for kw in _NAME_HINTS.get(role, ()) if kw in t)


def resolve_sheet(wb, role: str, *, exclude: set[str] | None = None,
                  default_first: bool = True) -> str | None:
    """Вернуть имя листа, наиболее подходящего под роль, или None.

    role: 'invoice' | 'packing' | 'details' | 'proforma'.
    exclude: имена листов, которые уже заняты другой ролью.
    default_first: если ничего не набрало очков — вернуть первый доступный лист
        (для одно-листовых книг), иначе None.

    Оценка: сигнатура колонок ×10 (сильный признак) + очки за имя. При равенстве
    выигрывает лист, идущий раньше в книге.
    """
    exclude = exclude or set()
    candidates = [s for s in wb.sheetnames if s not in exclude]
    if not candidates:
        return None

    best_name, best_score = None, 0
    for title in candidates:
        ws = wb[title]
        score = _column_signature_score(ws, role) * 10 + _name_score(title, role)
        if score > best_score:
            best_score, best_name = score, title

    if best_name is not None and best_score > 0:
        return best_name
    return candidates[0] if default_first else None
