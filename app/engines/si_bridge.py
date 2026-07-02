"""
si_bridge.py — мост canonical.Shipment → входные данные SI-движка.

generate_si_documents.read_source() умеет читать только ПОЛНОСТЬЮ заполненный
шаблон (обязательные листы Invoice List + Packing List + Details), иначе падает.
Чтобы транзитный документ и CMR строились ДАЖЕ когда целый шаблон не заполнен,
мы даём SI-движку те же кадры данных, что он ждёт от read_source, но собранные
из canonical-модели. Сам 44-КБ движок SI не трогаем — только кормим его.

Контракт колонок повторяет read_source():
  inv_df: N, Brand, Part_number, Description, Description_RU, Quantity,
          Weight_unit, Gross, Price, Amount, Origin, HSCode, Currency, BoxN
  pl_df : Box, Width, Height, Length, Weight, CBM
  details: {invoice_no, date, consignee, contract, incoterms}
  profile: ключ в generate_si_documents.PROFILES
"""
from __future__ import annotations

import pandas as pd

from .canonical import Shipment, _num


def _hs_str(v) -> str:
    """HS-код к строке без .0 (Int64-подобно, как в read_source)."""
    if v in (None, ""):
        return ""
    n = _num(v)
    if n is not None and float(n).is_integer():
        return str(int(n))
    return str(v).strip()


def build_si_inputs(sh: Shipment):
    """canonical → (inv_df, pl_df, details, profile) для generate_si_documents."""
    currency = (sh.requisites.currency or "USD").upper()

    inv_records = []
    for i, g in enumerate(sh.goods, start=1):
        inv_records.append({
            "N": i,
            "Brand": g.brand or "",
            "Part_number": g.part_number or "",
            "Description": g.description_en or "",
            "Description_RU": g.description_ru or "",
            "Quantity": _num(g.qty) or 0,
            "Weight_unit": _num(g.unit_weight),
            "Gross": g.effective_gross or 0,
            "Price": _num(g.unit_price),
            "Amount": g.effective_amount or 0,
            "Origin": g.coo or "",
            "HSCode": _hs_str(g.hs_code),
            "Currency": currency,
            "BoxN": g.box_number or "",
        })
    inv_df = pd.DataFrame(inv_records, columns=[
        "N", "Brand", "Part_number", "Description", "Description_RU",
        "Quantity", "Weight_unit", "Gross", "Price", "Amount", "Origin",
        "HSCode", "Currency", "BoxN"])

    pl_records = []
    for b in sh.boxes:
        pl_records.append({
            "Box": b.number,
            "Width": _num(b.width),
            "Height": _num(b.height),
            "Length": _num(b.length),
            "Weight": _num(b.gross_weight),
            "CBM": _num(b.cbm),
        })
    pl_df = pd.DataFrame(pl_records, columns=[
        "Box", "Width", "Height", "Length", "Weight", "CBM"])

    req = sh.requisites
    details = {
        "invoice_no": req.invoice_no or "",
        "date": req.date or "",
        "consignee": req.receiver_lines or [],
        "contract": req.contract or "",
        "incoterms": req.incoterms or "",
    }

    # Профиль влияет на дефолтную валюту/раскладку; отправитель/получатель всё
    # равно перекрываются из формы в service (si_data()).
    profile = "autoseller" if currency == "AED" else "fa_logistics"
    return inv_df, pl_df, details, profile
