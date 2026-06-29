"""
service.py — оркестратор генерации документов.

Связывает выбор пользователя (какие документы нужны + отправитель/получатель)
с движками из app/engines/. Каждый тип вывода обёрнут в try/except: ошибка
одного документа не ломает остальные — пользователь получает то, что удалось,
плюс понятные сообщения о том, что не получилось.

Движки используют импорты верхнего уровня (`from classifier import ...`,
`import hssc_txt_v3`), поэтому папку engines/ добавляем в sys.path как есть,
не переписывая их.
"""
from __future__ import annotations

import datetime
import re
import shutil
import sys
import traceback
from pathlib import Path

from .config import ENGINES_DIR

# --- подключаем движки (импорты верхнего уровня внутри них) ---
if str(ENGINES_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINES_DIR))

import openpyxl  # noqa: E402
import generate_si_documents as si  # noqa: E402
import hssc_txt_v3 as hssc_engine  # noqa: E402
import parties as parties_db  # noqa: E402
import fill_inv_pl  # noqa: E402

# Папка с готовыми шаблонами (INV/PL по типам шапки)
_APP_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = _APP_DIR / "templates_xlsx"
# Базовый шаблон один (OptiAuto); тип шапки меняет логотип в коде филлера.
_INV_BASE = TEMPLATES_DIR / "OptiAuto" / "INV_OptiAuto.xlsx"
_PL_BASE = TEMPLATES_DIR / "OptiAuto" / "PL_OptiAuto.xlsx"
INV_TEMPLATES = {"optiauto": _INV_BASE, "wens": _INV_BASE, "none": _INV_BASE}
PL_TEMPLATES = {"optiauto": _PL_BASE, "wens": _PL_BASE, "none": _PL_BASE}


def parties_for_ui() -> dict:
    """Реестр отправителей/получателей для выпадающих списков в UI.
    Каждый элемент: {key, name, block} — block это готовый текст для textarea.
    """
    def pack(table):
        items = []
        for key, p in table.items():
            block = "\n".join([p["name"], *p.get("lines", [])])
            items.append({"key": key, "name": p["name"], "block": block})
        return items
    return {
        "senders": pack(parties_db.SENDERS),
        "receivers": pack(parties_db.RECEIVERS),
    }


# Доступные типы вывода (ключ → подпись для UI). Сверху — три основных документа.
OUTPUT_TYPES = {
    "inv_ru":           "INV (RU) — инвойс с русскими наименованиями",
    "pl":               "PL — упаковочный лист",
    "transit_document": "Invoice Transit (INVOICE_TRANSIT)",
    "txt_hssc":         "txt HSSC codes (таможня Дубая)",
    "cmr_document":     "CMR document (таблица для накладной)",
    "whole_template":   "Весь шаблон (целиком)",
    "separate_sheets":  "Отдельные листы из шаблона",
}


def _safe(name: str) -> str:
    """Безопасное имя файла."""
    return re.sub(r"[^A-Za-z0-9_.\-]+", "_", str(name)).strip("_") or "file"


def _lines(text: str | None) -> list[str]:
    """Многострочное поле формы → список непустых строк."""
    if not text:
        return []
    return [ln.strip() for ln in text.replace("\r\n", "\n").split("\n") if ln.strip()]


class Result:
    def __init__(self):
        self.created: list[Path] = []   # успешно созданные файлы
        self.messages: list[str] = []   # предупреждения / ошибки для пользователя

    def add(self, path: Path):
        self.created.append(path)

    def warn(self, msg: str):
        self.messages.append(msg)


# --------------------------------------------------------------------------- #
# Подготовка данных для SI-движка (общая для transit + cmr)
# --------------------------------------------------------------------------- #

def _prepare_si(src_xlsx: Path, sender_lines: list[str], receiver_lines: list[str],
                invoice_no: str | None = None, date: str | None = None):
    """Читает источник, классифицирует, агрегирует, распределяет места.
    Возвращает (agg, details, profile, total_places, inv_no). Параметры
    отправителя/получателя и ручные номер/дата подставляются в профиль/детали.
    """
    inv_df, pl_df, details, profile = si.read_source(src_xlsx)

    # Ручной ввод перекрывает то, что прочитано из источника
    if invoice_no:
        details["invoice_no"] = invoice_no
    if date:
        details["date"] = date

    # Подстановка отправителя/получателя из формы (пусто → ничего не показываем)
    si.PROFILES[profile]["sender"] = sender_lines or []
    si.PROFILES[profile]["default_consignee"] = receiver_lines or []
    details["consignee"] = receiver_lines or details.get("consignee") or []

    enriched = si.enrich(inv_df)
    agg = si.aggregate(enriched)
    total_places = len(pl_df)
    agg = si.distribute_places(agg, total_places, enriched)

    # Масштабирование брутто к итогу из шапки, если он есть
    if details.get("_override_gross") and agg["gross"].sum() > 0:
        scale = details["_override_gross"] / agg["gross"].sum()
        agg["gross"] = agg["gross"] * scale

    inv_no = details.get("invoice_no") or src_xlsx.stem
    return agg, details, profile, total_places, inv_no


# --------------------------------------------------------------------------- #
# Точка входа
# --------------------------------------------------------------------------- #

def generate_documents(
    *,
    source_path: str | Path,
    selections: list[str],
    sender_text: str | None,
    receiver_text: str | None,
    out_dir: str | Path,
    invoice_no: str | None = None,
    date: str | None = None,
    header_type: str = "optiauto",
) -> Result:
    src = Path(source_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = _safe(src.stem)
    sender_lines = _lines(sender_text)
    receiver_lines = _lines(receiver_text)
    inv_label = _safe(invoice_no) if invoice_no else stem
    res = Result()

    # Нормализуем к .xlsx (для движков на openpyxl). .xls/.xlsb → конвертация
    # через LibreOffice; если его нет — pandas-движки (xlrd/pyxlsb) всё равно
    # читают транзит/CMR, но openpyxl-выводам нужен настоящий .xlsx.
    try:
        src_xlsx = si.to_xlsx_if_needed(src)
    except Exception:
        src_xlsx = src

    # Кэш агрегата SI, чтобы не считать дважды (transit + cmr)
    _si_cache = {}

    def si_data():
        if "v" not in _si_cache:
            _si_cache["v"] = _prepare_si(src_xlsx, sender_lines, receiver_lines,
                                         invoice_no, date)
        return _si_cache["v"]

    # ---- 0. INV (RU) — заполнение шаблона ----
    if "inv_ru" in selections:
        try:
            template = INV_TEMPLATES.get(header_type) or INV_TEMPLATES["optiauto"]
            if not Path(template).exists():
                raise RuntimeError(f"нет шаблона для шапки «{header_type}»")
            p = out_dir / f"INV_RU_{inv_label}.xlsx"
            rep = fill_inv_pl.fill_inv_ru(
                template_path=str(template),
                source_path=str(src_xlsx),
                output_path=str(p),
                invoice_no=invoice_no,
                date=date,
                receiver_lines=receiver_lines,
                sender_lines=sender_lines,
                header_type=header_type,
            )
            res.add(p)
            for note in rep.get("notes", []):
                res.warn(f"INV(RU): {note}")
        except Exception as e:
            res.warn(f"INV(RU): не удалось — {e}")
            traceback.print_exc()

    # ---- 0b. PL — упаковочный лист ----
    if "pl" in selections:
        try:
            template = PL_TEMPLATES.get(header_type) or PL_TEMPLATES["optiauto"]
            if not Path(template).exists():
                raise RuntimeError(f"нет PL-шаблона для шапки «{header_type}»")
            p = out_dir / f"PL_{inv_label}.xlsx"
            rep = fill_inv_pl.fill_pl(
                template_path=str(template),
                source_path=str(src_xlsx),
                output_path=str(p),
                invoice_no=invoice_no,
                date=date,
                receiver_lines=receiver_lines,
                sender_lines=sender_lines,
                header_type=header_type,
            )
            res.add(p)
            for note in rep.get("notes", []):
                res.warn(f"PL: {note}")
        except Exception as e:
            res.warn(f"PL: не удалось — {e}")
            traceback.print_exc()

    # ---- 1. Весь шаблон ----
    if "whole_template" in selections:
        try:
            if src_xlsx.suffix.lower() == ".xlsx":
                dst = out_dir / f"{stem}.xlsx"
            else:
                dst = out_dir / src.name   # отдаём как есть
            shutil.copy(src_xlsx, dst)
            res.add(dst)
        except Exception as e:
            res.warn(f"Весь шаблон: не удалось — {e}")

    # ---- 2. Отдельные листы ----
    if "separate_sheets" in selections:
        try:
            if src_xlsx.suffix.lower() != ".xlsx":
                raise RuntimeError(
                    "нужен .xlsx (для .xls/.xlsb установите LibreOffice)")
            probe = openpyxl.load_workbook(src_xlsx, read_only=True)
            names = list(probe.sheetnames)
            probe.close()
            for name in names:
                wb = openpyxl.load_workbook(src_xlsx)
                for other in list(wb.sheetnames):
                    if other != name:
                        del wb[other]
                p = out_dir / f"{stem}__{_safe(name)}.xlsx"
                wb.save(p)
                wb.close()
                res.add(p)
            if not names:
                res.warn("Отдельные листы: в файле нет листов.")
        except Exception as e:
            res.warn(f"Отдельные листы: не удалось — {e}")

    # ---- 3. Транзитный документ ----
    if "transit_document" in selections:
        try:
            agg, details, profile, total_places, inv_no = si_data()
            p = out_dir / f"INVOICE_TRANSIT_{_safe(inv_no)}.xlsx"
            si.build_invoice_transit(
                agg, details, profile, p,
                blank=False, display_invoice_no=inv_no)
            res.add(p)
        except Exception as e:
            res.warn(f"Транзитный документ: не удалось — {e}")
            traceback.print_exc()

    # ---- 4. CMR ----
    if "cmr_document" in selections:
        try:
            agg, details, profile, total_places, inv_no = si_data()
            p = out_dir / f"CMR_{_safe(inv_no)}.xlsx"
            si.build_cmr_table(agg, details, profile, total_places, p, blank=False)
            res.add(p)
        except Exception as e:
            res.warn(f"CMR: не удалось — {e}")
            traceback.print_exc()

    # ---- 5. TXT HSSC ----
    if "txt_hssc" in selections:
        try:
            if src_xlsx.suffix.lower() != ".xlsx":
                raise RuntimeError(
                    "нужен .xlsx (для .xls/.xlsb установите LibreOffice)")
            wb = openpyxl.load_workbook(src_xlsx, data_only=True)
            iso2_to_name = {iso2: full
                            for full, iso2 in hssc_engine.COUNTRY_NAME_TO_ISO2.items()}
            sheet = next((c for c in ("HS CODE SUM", "HSSC") if c in wb.sheetnames), None)
            if sheet:
                agg_rows = hssc_engine.aggregate_hs_code_sum(wb[sheet], iso2_to_name)
            else:
                inv_name = next((c for c in ("INV", "INVOICE") if c in wb.sheetnames), None)
                if not inv_name:
                    raise RuntimeError("нет листа HS CODE SUM / HSSC / INV")
                agg_rows = hssc_engine.aggregate_inv(wb[inv_name])
            wb.close()
            if not agg_rows:
                raise RuntimeError("не найдено строк с HS-кодами")
            company = sender_lines[0] if sender_lines else "WENS LOGISTICS DWC-LLC"
            txt = hssc_engine.generate_txt(
                agg_rows,
                decl_no=stem,
                decl_date=datetime.date.today().strftime("%Y-%m-%d"),
                pages=1,
                company=company,
                supplier_code=1,
                incoterm="EXW",
                hs_units={},   # TODO Спринт 5: реальная база dt_hscodes (CD_BLANK_NEW.xlsx)
            )
            p = out_dir / f"{stem}_HSSC.txt"
            p.write_bytes(txt.encode("utf-8"))
            res.add(p)
        except Exception as e:
            res.warn(f"TXT HSSC: не удалось — {e}")

    if not res.created and not res.messages:
        res.warn("Не выбран ни один тип документа.")
    return res
