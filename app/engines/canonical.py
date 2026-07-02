"""
canonical.py — единый внутренний формат («canonical model»).

Что бы ни пришло на вход (1 файл Combined, 2 файла INV+PL, 3 файла с проформой,
разные форматы xls/xlsb/xlsx и разные названия листов/колонок) — адаптеры из
app/engines/sources/ приводят это к ОДНОЙ структуре `Shipment`. Все заполнители
и сборщики документов (INV/PL/Transit/CMR/TXT) читают ТОЛЬКО эту структуру и не
знают, как выглядел исходник.

См. АРХИТЕКТУРА_исходников.md §1.

Правило проекта: перевод НЕ выдумываем. Если русского наименования нет в
источнике — `description_ru` остаётся пустым (None), а не заполняется догадкой.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Optional


def _num(v) -> Optional[float]:
    """Мягкое приведение к числу: '1 234,50' → 1234.5, мусор → None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "")
    # запятая как десятичный разделитель, если нет точки
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


@dataclass
class Good:
    """Одна товарная позиция инвойса."""
    brand: Optional[str] = None
    part_number: Optional[str] = None
    description_en: Optional[str] = None
    description_ru: Optional[str] = None
    qty: Optional[float] = None
    unit_weight: Optional[float] = None       # вес за единицу
    gross_weight: Optional[float] = None      # общий вес позиции (если дан явно)
    unit_price: Optional[float] = None
    amount: Optional[float] = None            # сумма по позиции
    coo: Optional[str] = None                 # страна происхождения (Origin)
    hs_code: Optional[str] = None             # ТНВЭД / HS
    box_number: Optional[str] = None          # к какой коробке относится (если есть)
    incoming_decl: Optional[str] = None

    @property
    def effective_amount(self) -> Optional[float]:
        """Сумма: явная, иначе qty×unit_price (если оба числа)."""
        if self.amount not in (None, ""):
            return _num(self.amount)
        q, p = _num(self.qty), _num(self.unit_price)
        return q * p if q is not None and p is not None else None

    @property
    def effective_gross(self) -> Optional[float]:
        """Общий вес: явный gross, иначе unit_weight×qty."""
        if self.gross_weight not in (None, ""):
            return _num(self.gross_weight)
        q, w = _num(self.qty), _num(self.unit_weight)
        return q * w if q is not None and w is not None else None


@dataclass
class Box:
    """Одно грузовое место (коробка) из упаковочного листа."""
    number: Optional[str] = None
    width: Optional[float] = None
    height: Optional[float] = None
    length: Optional[float] = None
    qty: Optional[float] = None               # мест в этой строке (обычно 1)
    gross_weight: Optional[float] = None
    cbm: Optional[float] = None
    actual_weight: Optional[float] = None


@dataclass
class Requisites:
    """Реквизиты отгрузки (шапка документов)."""
    invoice_no: Optional[str] = None
    date: Optional[str] = None
    currency: Optional[str] = None
    contract: Optional[str] = None
    incoterms: Optional[str] = None
    sender_lines: list[str] = field(default_factory=list)     # адрес отправителя (по строкам)
    receiver_lines: list[str] = field(default_factory=list)   # адрес получателя (по строкам)
    sender_key: Optional[str] = None          # ключ из parties.SENDERS (напр. 'optiauto')
    receiver_key: Optional[str] = None         # ключ из parties.RECEIVERS

    def merged(self, *, invoice_no=None, date=None, currency=None,
               sender_lines=None, receiver_lines=None,
               sender_key=None, receiver_key=None) -> "Requisites":
        """Копия с перекрытием непустыми значениями (ручной ввод > источник)."""
        return replace(
            self,
            invoice_no=invoice_no or self.invoice_no,
            date=date or self.date,
            currency=currency or self.currency,
            sender_lines=sender_lines if sender_lines else self.sender_lines,
            receiver_lines=receiver_lines if receiver_lines else self.receiver_lines,
            sender_key=sender_key or self.sender_key,
            receiver_key=receiver_key or self.receiver_key,
        )


@dataclass
class Shipment:
    """Единый внутренний формат: товары + коробки + реквизиты."""
    goods: list[Good] = field(default_factory=list)
    boxes: list[Box] = field(default_factory=list)
    requisites: Requisites = field(default_factory=Requisites)
    source_format: Optional[str] = None       # имя сработавшего адаптера (для UI/лога)

    # --- сводки для предпросмотра и распределения мест ---
    @property
    def total_rows(self) -> int:
        return len(self.goods)

    @property
    def total_places(self) -> int:
        """Число грузовых мест = число коробок в упаковочном листе."""
        return len(self.boxes)

    @property
    def total_qty(self) -> float:
        return sum(_num(g.qty) or 0 for g in self.goods)

    @property
    def total_gross(self) -> float:
        return sum(g.effective_gross or 0 for g in self.goods)

    def summary(self) -> dict:
        """Короткая сводка распознанного исходника (для показа перед генерацией)."""
        return {
            "format": self.source_format,
            "rows": self.total_rows,
            "places": self.total_places,
            "qty": self.total_qty,
            "gross": self.total_gross,
            "invoice_no": self.requisites.invoice_no,
            "currency": self.requisites.currency,
        }
