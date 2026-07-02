"""
registry.py — реестр адаптеров. Пробуем по порядку, первый подошедший побеждает.

Фаза 1: только Combined (1 файл). Фаза 2 добавит InvPlAdapter (2 файла) и
NovaAdapter (3 файла с проформой) — они встанут ВЫШЕ Combined, т.к. более
специфичны по числу файлов.
"""
from __future__ import annotations

from pathlib import Path

from ..canonical import Shipment
from .combined import CombinedAdapter

# Порядок важен: более специфичные адаптеры — раньше.
ADAPTERS = [
    CombinedAdapter(),
]


def pick(paths: list[Path]):
    """Вернуть первый адаптер, подходящий под набор файлов, или None."""
    for adapter in ADAPTERS:
        if adapter.matches(paths):
            return adapter
    return None


def read_shipment(paths: list[str] | list[Path]) -> Shipment:
    """Распознать исходник(и) и прочитать в единый формат Shipment."""
    paths = [Path(p) for p in paths]
    adapter = pick(paths)
    if adapter is None:
        names = ", ".join(p.name for p in paths) or "—"
        raise RuntimeError(f"Не найден адаптер для набора файлов: {names}")
    return adapter.read(paths)
