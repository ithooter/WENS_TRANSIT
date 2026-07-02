"""
template_select.py — выбор шаблона по ОТПРАВИТЕЛЮ.

Правило (по договорённости):
  • отправитель OptiAuto            → шаблон OptiAuto (с логотипом);
  • отправитель WENS                → шаблон WENS (пока не готов → ведём как Базовый);
  • любой другой / свободный ввод    → Базовый = тот же OptiAuto, но БЕЗ логотипа
                                       (сверху пишется адрес отправителя).

Раньше выбор шапки был ручным радио-полем в форме (header_type). Теперь тип
шапки выводится из отправителя, а поле в UI убирается.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .parties import ALIASES

_TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates_xlsx"
OPTIAUTO_WHOLE = _TEMPLATES_DIR / "OptiAuto" / "OptiAuto_whole.xlsx"


@dataclass
class TemplateChoice:
    key: str                     # 'optiauto' | 'basic' | 'wens'
    whole_template: Path         # единый workbook с листами INV + PL
    remove_logo: bool            # писать адрес отправителя поверх шапки (Базовый)
    sender_key: str | None
    notes: list[str] = field(default_factory=list)


def select_template(sender_key: str | None) -> TemplateChoice:
    """Определить шаблон и режим шапки по ключу отправителя."""
    key = (sender_key or "").strip().lower().replace(" ", "_")
    key = ALIASES.get(key, key)

    if key == "optiauto":
        return TemplateChoice("optiauto", OPTIAUTO_WHOLE,
                              remove_logo=False, sender_key="optiauto")
    if key == "wens":
        return TemplateChoice(
            "wens", OPTIAUTO_WHOLE, remove_logo=True, sender_key="wens",
            notes=["Шаблон WENS ещё не готов — используем Базовый "
                   "(OptiAuto без логотипа)."])
    # Базовый — для всех остальных отправителей и свободного ввода.
    return TemplateChoice("basic", OPTIAUTO_WHOLE,
                          remove_logo=True, sender_key=key or None)
