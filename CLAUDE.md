# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Flask web app that turns a customer's Excel "исходник" (source) into transit
customs documents (INV(RU), PL, Invoice Transit, CMR, HSSC TXT). User logs in,
uploads a file, picks which documents they need + sender/receiver, downloads the
result. UI and domain terms are Russian; code identifiers are English.

## Commands

```bash
# Setup
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt          # runtime
pip install -r requirements-dev.txt      # + pytest (tests)

# Run (http://127.0.0.1:8000, debug reloader on)
python3 run.py                           # or: flask --app run run --debug

# Tests
.venv/bin/python -m pytest tests/ -q
.venv/bin/python -m pytest tests/test_pipeline.py::test_expansion_keeps_totals_consistent -q  # single

# .xls / .xlsb support needs LibreOffice (converts to .xlsx via to_xlsx_if_needed)
brew install --cask libreoffice
```

There is no lint/build step. The dev server auto-reloads on file changes.

## Architecture: the two-layer pipeline

The core design is **adapters → canonical model → fillers/builders**. Everything
funnels through one internal format so a new source layout means "add one
adapter" and a new output means "add one filler" — neither touches the other.

```
files (1–3, xls/xlsb/xlsx)
  └─ si.to_xlsx_if_needed (LibreOffice)         normalize to .xlsx
  └─ engines/sources/registry.read_shipment()   pick adapter, fuzzy-resolve sheets
        └─ canonical.Shipment(goods[], boxes[], requisites)
              ├─ fillers/optiauto.fill_optiauto_whole()  → fills ONE workbook (INV+PL sheets)
              │      service._extract_sheet(...)          → INV(RU) / PL are that workbook
              │                                             with the other sheet deleted
              └─ si_bridge.build_si_inputs()  → (inv_df, pl_df, details, profile)
                     └─ generate_si_documents.enrich/aggregate/distribute_places
                        → build_invoice_transit / build_cmr_table
```

[app/service.py](app/service.py) is the orchestrator: it reads the Shipment
**once** (lazily cached), then each output type runs in its own try/except so one
failure doesn't block the others — warnings surface to the user via `Result`.

**Key consequence of the "Весь шаблон = filled" decision:** INV(RU), PL and
"whole template" are all ONE fill of `OptiAuto_whole.xlsx`; INV(RU)/PL are just
that filled workbook with the other sheet stripped. Do not re-introduce separate
per-document fill code.

## `app/engines/` has TWO import worlds — know which you're in

- **Legacy engines** (`generate_si_documents`, `hssc_txt_v3`, `classifier`,
  `company_profiles`, `parties`, `fill_optiauto`, `fill_nova`, `fill_inv_pl`) use
  **top-level imports** of each other (`from classifier import ...`). These
  resolve because [app/engines/__init__.py](app/engines/__init__.py) inserts the
  engines dir into `sys.path` on package import (and `service.py` does too).
- **New pipeline modules** (`canonical`, `sheet_writer`, `sheet_resolver`,
  `sources/`, `fillers/`, `si_bridge`, `template_select`) are a normal
  `app.engines.*` package using **relative imports** (`from ..canonical import`).

When adding code, match the surrounding module's style. New work should live in
the package world and only reach into legacy modules for specific helpers.

## Non-obvious things that will bite you

- **The SI engine's own `read_source()` is bypassed in the web flow.** It hard-
  requires a fully-filled template (Invoice List + Packing List + **Details**
  sheets) and raises otherwise. Transit/CMR are instead fed from canonical via
  [si_bridge.py](app/engines/si_bridge.py) — whose DataFrame column contract must
  exactly mirror what `read_source` produced (documented at the top of that file).
- **Template geometry is data, not code.** `SheetGeometry` in
  [fillers/optiauto.py](app/engines/fillers/optiauto.py) (`INV_GEOM`/`PL_GEOM`)
  describes the real `templates_xlsx/OptiAuto/*.xlsx` layout (INV data rows
  23–185, totals 186/187/188; PL data 21–41, totals C43–C45). The generic
  [sheet_writer.py](app/engines/sheet_writer.py) inserts rows for >capacity input,
  copies the first data row's style, **shifts formulas with openpyxl `Translator`**
  (not string replace), and each geometry's `on_expand` callback rewrites the
  totals/header refs. If you change a template file, update its geometry.
- **`OptiAuto_whole.xlsx` is a built artifact**, merged from the two single-sheet
  `INV_OptiAuto.xlsx` + `PL_OptiAuto.xlsx` templates (both have 0 embedded images).
- **Sender drives the template.** [template_select.py](app/engines/template_select.py):
  `optiauto` → OptiAuto (logo kept); anything else / free text → Basic (same file,
  logo removed, sender address written on top); `wens` → deferred (falls back to
  Basic). The old `header_type` radio is gone; the sender `<select>` submits
  `sender_key`.
- **Never guess a translation.** If a good has no Russian description in the
  source, `description_ru` stays empty — do not fill a fallback. This is a hard
  project rule.
- **Broken `[1]` VLOOKUP refs are expected.** The INV template's helper columns
  (T,U,V,W,Z,AB) reference an external `[1]` workbook that isn't shipped; they're
  cloned as-is onto inserted rows and left to resolve (or show cached values).
- **`classifier.py` rules are order-sensitive** (first regex match wins; oil
  detection by HS `2710` prefix takes priority). It maps a row to a TNVED group
  used for Transit aggregation.
- **Row-existence uses `not in (None, "")`, and money/weight go through
  `canonical._num`** (handles `"1 234,50"` style). A qty/price of `0` is valid —
  don't treat it as missing.

## State & data

SQLite (no ORM) via [app/db.py](app/db.py): `users` + `jobs` tables, created on
startup. Uploaded sources and generated outputs live under `instance/`
(gitignored), namespaced by user id and a per-job UUID. `SECRET_KEY` defaults to
a dev value; set it via env for anything real.

> The README.md and ROADMAP.md predate the Sprint-2 pipeline refactor and are
> partly stale (they still describe "Весь шаблон" as a raw copy, an "Отдельные
> листы" option, and `classifier.py` as a stub). Trust the code and this file.
> `tests/` covers the canonical pipeline; the next planned work (Phase 2) is
> multi-file upload + the INV+PL and NOVA/proforma adapters in `sources/`.
