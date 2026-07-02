"""Общие фикстуры: строим синтетический Combined-исходник в памяти (без зависимости
от реальных файлов из instance/, которые в .gitignore)."""
import openpyxl
import pytest

INV_HEADERS = ["Brand", "Part number", "Description", "Description Rus",
               "Quantity", "Weight", "Price AED", "Total AED", "Origin", "HSCode"]
PL_HEADERS = ["Box N", "Width", "Height", "Length", "Weight", "CBM"]


def _build_combined(path, n_goods, n_boxes, inv_sheet="Invoice List",
                    pl_sheet="Packing List"):
    wb = openpyxl.Workbook()
    inv = wb.active
    inv.title = inv_sheet
    inv.append(INV_HEADERS)
    for i in range(1, n_goods + 1):
        inv.append([f"BRAND{i}", f"PN{i:04d}", f"Item {i} EN",
                    (f"Товар {i}" if i % 2 == 0 else None),   # часть без RU-перевода
                    i, 1.5, 10.0, 10.0 * i, "UNITED ARAB EMIRATES", 8708999709])
    pl = wb.create_sheet(pl_sheet)
    pl.append(PL_HEADERS)
    for j in range(1, n_boxes + 1):
        pl.append([100000 + j, 40, 30, 20, 12.0, 0.024])
    wb.save(path)
    wb.close()


@pytest.fixture
def combined_small(tmp_path):
    """5 товаров, 3 коробки — влезает в шаблон без расширения."""
    p = tmp_path / "combined_small.xlsx"
    _build_combined(p, 5, 3)
    return p


@pytest.fixture
def combined_large(tmp_path):
    """200 товаров, 30 коробок — БОЛЬШЕ вместимости шаблона (163/21) → расширение."""
    p = tmp_path / "combined_large.xlsx"
    _build_combined(p, 200, 30)
    return p


@pytest.fixture
def combined_renamed(tmp_path):
    """Листы без узнаваемых имён — резолвер должен опознать по колонкам."""
    p = tmp_path / "combined_renamed.xlsx"
    _build_combined(p, 5, 3, inv_sheet="Processed Data", pl_sheet="XYZ Random")
    return p
