"""
Company profiles — classification of exporters/logistics companies.

Each company that ships through the customs pipeline has its own "profile"
describing the bits that differ between companies:
  - the company name written into the IH header,
  - the declared currency (USD / AED / ...),
  - the supplier code (IH field [6]),
  - the default incoterm,
  - where to find/extract metadata in the source workbook.

The customs TXT *format* itself (15-field IH, 18-field ID, YY-MM-DD date,
LF line endings, etc.) is shared across companies — it's the Dubai customs
spec — and lives in `hssc_txt_v3.py`. Only the values differ per company.

Add a new company by adding one PROFILES entry. Nothing else changes.

Calibration status:
  WENS     — ✅ TXT byte-for-byte verified (TPL2605012.txt, et al.)
  OPTIAUTO — ⚠️ profile built from base template OPTIAUTO_APMG_00_2025.xlsb.
             No reference TXT yet → currency/supplier_code/company_name are
             best-effort defaults. Supply an OptiAuto reference TXT to lock.
"""

from __future__ import annotations
from dataclasses import dataclass, field


@dataclass
class CompanyProfile:
    key: str                       # canonical key, lowercase: 'wens', 'optiauto'
    cli_flags: tuple[str, ...]     # accepted CLI flags, e.g. ('--Wens', '--wens')
    company_name: str              # name written into IH header [5]
    currency: str                  # IH header [8]: 'USD' / 'AED'
    supplier_code: int             # IH header [6]
    incoterm_default: str          # IH header [10]
    mode: str = "txt"              # 'txt' = generate customs TXT
    #                                'fill_template' = fill INV+PL, keep design
    consignee: str = ""            # informational only (BILL TO)
    # Source-workbook hints (most are auto-detected, these are fallbacks):
    hssc_sheet: str = "HSSC"
    decl_no_label: str = "TAX Invoice Number"
    notes: str = ""
    calibrated: bool = False       # True once a reference TXT confirms format


# --------------------------------------------------------------------------- #
# Registry
# --------------------------------------------------------------------------- #

PROFILES: dict[str, CompanyProfile] = {
    "wens": CompanyProfile(
        key="wens",
        cli_flags=("--Wens", "--wens", "--WENS"),
        company_name="WENS LOGISTICS DWC-LLC",
        currency="USD",
        supplier_code=1,
        incoterm_default="EXW",
        consignee="(various)",
        notes=(
            "Includes legacy OMEGAPART DWC LLC (old name) and the TOP PARTS "
            "supplier variant. TXT format byte-for-byte verified."
        ),
        calibrated=True,
    ),
    "optiauto": CompanyProfile(
        key="optiauto",
        cli_flags=("--OptiAuto", "--optiauto", "--OPTIAUTO", "--Optiauto"),
        company_name="OptiAuto FZE",       # from OTH sheet of base template
        currency="AED",                    # source INV declares AED (read from source)
        supplier_code=1,                   # n/a for fill mode
        incoterm_default="EXW",
        mode="fill_template",              # OptiAuto: fill INV+PL, NOT generate TXT
        consignee="LLC «Autoseller»",
        notes=(
            "Base template OPTIAUTO_APMG_00_2025.xlsb. Output = filled INV + PL "
            "with original design preserved (header, addresses, red headers, "
            "totals formulas). No TXT. Currency taken from source."
        ),
        calibrated=False,
    ),
    "carbotec": CompanyProfile(
        key="carbotec",
        cli_flags=("--Carbotec", "--carbotec", "--CARBOTEC", "--CarboTec"),
        company_name="CarboTec Trading FZ-LLC",
        currency="USD",                    # source declares USD
        supplier_code=1,                   # n/a for fill mode
        incoterm_default="EXW",
        mode="fill_template",
        consignee="L.L.C. 100PART",
        notes=(
            "Same template as OptiAuto with logos removed and addresses "
            "replaced (sender CarboTec Trading FZ-LLC; consignee L.L.C. 100PART). "
            "Currency USD. Source sheets: 'Processed Data' (INV), 'Packing List'. "
            "Use CARBOTEC_template.xlsx as the --template."
        ),
        calibrated=False,
    ),
    "llcnova": CompanyProfile(
        key="llcnova",
        cli_flags=("--LLCNova", "--llcnova", "--LLCNOVA", "--Nova", "--nova"),
        company_name="ELITEGOODS TRADING LLC-FZ",   # sender/exporter (FZ)
        currency="USD",                    # source declares USD
        supplier_code=1,                   # n/a for fill mode
        incoterm_default="EXW",            # EXW Dubai (placed in terms block)
        mode="fill_template",
        consignee="LLC NOVA",
        notes=(
            "Same template as OptiAuto with logos removed and addresses "
            "replaced (sender ELITEGOODS TRADING LLC-FZ; consignee LLC NOVA, "
            "incl. SALES CONTRACT line + EXW Dubai). Currency USD, HS 8-digit. "
            "INV source = commercial invoice (Sheet_1), 283 lines; RU description "
            "+ unit weight enriched from the proforma Sheet1 by article (ITEM "
            "CODE↔Артикул) via inv_enrichment. PL = one row per (item×case); "
            "dedup by CASE NO (pl_dedup_field='pkg_number') → 104 boxes. "
            "Use NOVA_template.xlsx as the --template."
        ),
        calibrated=False,
    ),
}


# Flag → profile-key lookup, built from each profile's cli_flags.
_FLAG_TO_KEY: dict[str, str] = {}
for _key, _prof in PROFILES.items():
    for _flag in _prof.cli_flags:
        _FLAG_TO_KEY[_flag.lower()] = _key
    _FLAG_TO_KEY[f"--{_key}"] = _key  # always allow --<key>


def resolve_profile(flag_or_key: str) -> CompanyProfile:
    """Resolve a CLI flag ('--OptiAuto') or bare key ('optiauto') to a profile.

    Raises KeyError with a helpful message if unknown.
    """
    s = flag_or_key.strip()
    key = _FLAG_TO_KEY.get(s.lower())
    if key is None:
        # also accept bare key without dashes
        key = s.lower().lstrip("-")
    if key not in PROFILES:
        valid = ", ".join(sorted({f for p in PROFILES.values() for f in p.cli_flags}))
        raise KeyError(
            f"Unknown company '{flag_or_key}'. Valid flags: {valid}"
        )
    return PROFILES[key]


def list_companies() -> str:
    """Human-readable listing of all registered companies."""
    lines = []
    for key, p in PROFILES.items():
        status = "✅ calibrated" if p.calibrated else "⚠️  uncalibrated"
        lines.append(
            f"  {p.cli_flags[0]:<14} {p.company_name:<35} "
            f"{p.currency}  {status}"
        )
    return "\n".join(lines)
