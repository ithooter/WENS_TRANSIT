"""
parties.py — address & requisites classifier for the customs pipeline.

A small in-code "database" of every sender (exporter) and receiver (consignee)
the pipeline has handled, plus a helper to stamp a sender/receiver pair into an
INV+PL template (the OptiAuto base layout). Keeps all addresses in ONE place so
a new shipment is just `apply_parties(wb, sender="optiauto", receiver="master_parts")`
instead of re-typing address blocks each time.

Each entry:
  name   — display name (line 1 of the block)
  lines  — the remaining address / requisites lines, in order
  meta   — free-form dict (inn, ogrn, contract, incoterm, currency, tel, email…)

Template geometry (OptiAuto base, shared by all fill clients):
  SENDER block  : INV rows 2-7, column B (logo sits here when kept)
  BILL TO block : INV rows 15-20, column B   (consignee)
  SHIP TO block : INV rows 15-20, column E   (mirror of BILL TO; E15 == B15)
"""

from __future__ import annotations

# --------------------------------------------------------------------------- #
# SENDERS (exporters)
# --------------------------------------------------------------------------- #
SENDERS: dict[str, dict] = {
    "optiauto": {
        "name": "OptiAuto FZE",
        "lines": [],                       # logo block kept; no text override
        "meta": {"currency": "AED", "keep_logo": True},
    },
    "carbotec": {
        "name": "CarboTec Trading FZ-LLC Business Center RAKEZ",
        "lines": [
            "Ras Al Khaiman, United Arab Emirates",
            "UAE, Dubai, Motor City, Detroit house, office 907",
            "Tel.: +971553771122,",
            "email: office@carbotecfze.com",
        ],
        "meta": {"currency": "USD"},
    },
    "elitegoods": {
        "name": "ELITEGOODS TRADING LLC-FZ",
        "lines": [
            "Meydan Grandstand, 6th floor, Meydan Road,",
            "Nad Al Sheba, Dubai, UAE",
        ],
        "meta": {"currency": "USD"},
    },
    "proconstruct": {
        "name": "ProConstruct Limited",
        "lines": [
            "UNIT A7 12/F ASTORIA BUILDING 34 ASHLEY ROAD",
            "TSIM SHA TSUI KL, HONG KONG",
        ],
        "meta": {"currency": "USD"},
    },
}

# --------------------------------------------------------------------------- #
# RECEIVERS (consignees)
# --------------------------------------------------------------------------- #
RECEIVERS: dict[str, dict] = {
    # ---- new (OptiAuto sender) ----
    "apm_parts": {
        "name": "ООО «APM-PARTS»",
        "lines": [
            "127521, 13c5, 17-proezd Maryinoy Roschi,",
            "Moscow, Russia",
            "INN 7724424423; OGRN 5177746233132",
            "Attn.: +7 926 020-2005",
            "om@log-pro.com",
        ],
        "meta": {"inn": "7724424423", "ogrn": "5177746233132"},
    },
    "ip_vlasov": {
        "name": "IP Alexey Vladimirovich Vlasov",
        "lines": [
            "Tukhachevskogo street 16-1-56",
            "123423, Moscow, Russia.",
            "Contract No and Date: 26032025 From 26.03.2025",
        ],
        "meta": {"contract": "26032025", "contract_date": "26.03.2025"},
    },
    "master_parts": {
        "name": "“MASTER PARTS” LLC",
        "lines": [
            "47 Aviastroiteley Street, Office 3,",
            "Novosibirsk Region, Novosibirsk, Russia",
            "INN: 5404005910 PPC: 541001001",
            "Phone: +7-913-000-7-000",
        ],
        "meta": {"inn": "5404005910", "ppc": "541001001"},
    },
    "log_pro": {
        "name": "LOG-PRO LLC",
        "lines": [
            "Apt 12, 18 Moskovskaya str., Lubertsy, 140011",
            "RUSSIAN FEDERATION",
            "+7 926 276 53 23",
        ],
        "meta": {},
    },
    # ---- existing ----
    "llc_nova": {
        "name": "LLC NOVA",
        "lines": [
            "109147, Moscow, inner city, Tagansky municipal district,",
            "Marxistskaya street, building 34, office 8, room 1/6",
            "SALES CONTRACT No.27/01/2026-1, 27 January 2026",
            "EXW Dubai, UAE",
        ],
        "meta": {"incoterm": "EXW", "currency": "USD"},
    },
    "auto_euro": {
        "name": "«AUTO-EURO» Joint-stock company",
        "lines": [
            "108810, Moscow, vn.ter.g. Municipal district Vnukovo,",
            "v. Sovhoza Kryokshino, str. Ozernaya, b. 5 b. 1, R. 27",
            "CONTRACT OF SALE No. DXB-DAP-AE-20032026 dd March 20, 2026",
            "DAP, village of Otyakovo, Mozhaysky Urban District, Moscow Region",
        ],
        "meta": {"incoterm": "DAP", "currency": "USD"},
    },
    "fa_logistic": {
        "name": "F.A. Logistic",
        "lines": [
            "4, Solnechnogorsky proezd,",
            "125413, Moscow, Russian Federation",
            "Tel.: + 7 (495) 789-80-00",
        ],
        "meta": {},
    },
    "autoseller": {
        "name": "LLC «Autoseller»",
        "lines": [],   # full block already baked into the OPTIAUTO template
        "meta": {"currency": "AED"},
    },
}

# Aliases so a caller can look a party up by a human label too.
ALIASES = {
    "apm-parts": "apm_parts", "apm": "apm_parts",
    "vlasov": "ip_vlasov", "ip": "ip_vlasov",
    "master": "master_parts", "master parts": "master_parts",
    "logpro": "log_pro", "log-pro": "log_pro", "log pro": "log_pro",
    "nova": "llc_nova", "autoeuro": "auto_euro", "auto-euro": "auto_euro",
    "fa": "fa_logistic", "falogistic": "fa_logistic",
}

SENDER_ROWS = range(2, 8)      # INV rows that hold the sender block (col B)
RECV_ROWS = range(15, 21)      # INV rows that hold BILL TO (col B) / SHIP TO (col E)


def _resolve(key: str, table: dict) -> dict:
    k = key.strip().lower().replace(" ", "_")
    if k in table:
        return table[k]
    if k in ALIASES and ALIASES[k] in table:
        return table[ALIASES[k]]
    raise KeyError(f"party '{key}' not found in {list(table)}")


def get_sender(key: str) -> dict:
    return _resolve(key, SENDERS)


def get_receiver(key: str) -> dict:
    return _resolve(key, RECEIVERS)


def _block_lines(party: dict) -> list[str]:
    return [party["name"], *party.get("lines", [])]


def apply_parties(wb, *, sender: str | None = None, receiver: str | None = None,
                  remove_logo: bool = False) -> dict:
    """Stamp a sender and/or receiver address block into an INV+PL workbook.

    • sender   — key into SENDERS. If the sender's meta has keep_logo=True and
                 its lines are empty (OptiAuto), the logo/sender block is left
                 untouched. Otherwise the sender block (rows 2-7, col B) is set.
    • receiver — key into RECEIVERS; fills BILL TO (col B) + SHIP TO (col E).
    • remove_logo — clear INV & PL images (used by non-OptiAuto senders).

    Returns the meta dict of {'sender': ..., 'receiver': ...} so the caller can
    pick up currency / incoterm defaults.
    """
    inv = wb["INV"]
    out = {}

    if remove_logo:
        for sh in ("INV", "PL"):
            if sh in wb.sheetnames and hasattr(wb[sh], "_images"):
                wb[sh]._images = []

    if sender is not None:
        s = get_sender(sender)
        out["sender"] = s.get("meta", {})
        if not (s.get("meta", {}).get("keep_logo") and not s["lines"]):
            lines = _block_lines(s)
            for i, r in enumerate(SENDER_ROWS):
                inv.cell(row=r, column=2).value = lines[i] if i < len(lines) else None

    if receiver is not None:
        rv = get_receiver(receiver)
        out["receiver"] = rv.get("meta", {})
        lines = _block_lines(rv)
        for i, r in enumerate(RECV_ROWS):
            val = lines[i] if i < len(lines) else None
            inv.cell(row=r, column=2).value = val                       # BILL TO
            inv.cell(row=r, column=5).value = "=B15" if i == 0 else val  # SHIP TO
        # leftover PL address line points at INV B19/E19 (resolved on literalize)
        if "PL" in wb.sheetnames:
            wb["PL"].cell(row=17, column=2).value = "=INV!B19"
            wb["PL"].cell(row=17, column=7).value = "=INV!E19"

    return out


def list_parties() -> str:
    s = "SENDERS:\n" + "\n".join(f"  {k}: {v['name']}" for k, v in SENDERS.items())
    r = "RECEIVERS:\n" + "\n".join(f"  {k}: {v['name']}" for k, v in RECEIVERS.items())
    return s + "\n" + r


if __name__ == "__main__":
    print(list_parties())
