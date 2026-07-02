"""
app.engines — движки документов.

Новые модули (canonical, sheet_writer, sources/, fillers/, si_bridge,
template_select, sheet_resolver) — это пакет `app.engines.*`.

Легаси-движки (generate_si_documents, hssc_txt_v3, classifier, company_profiles,
parties, fill_inv_pl, fill_optiauto, fill_nova) написаны на импортах ВЕРХНЕГО
уровня (`from classifier import ...`, `from hssc_txt_v3 import ...`). Чтобы они
резолвились и когда пакет импортируется как `app.engines`, добавляем папку
engines в sys.path здесь — один раз, при импорте пакета.
"""
import sys
from pathlib import Path

_HERE = str(Path(__file__).resolve().parent)
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
