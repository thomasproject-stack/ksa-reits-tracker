# KSA REITs Tracker

**A free-first pipeline that snapshots every Saudi REIT on Tadawul straight from Yahoo Finance's public API and renders a styled, formula-driven Excel report — bilingual (EN/FR), with a methodology sheet that cites every source.** No paid data feed, no API key: the whole thing runs on three small scripts.

---

## Why it's interesting

Most "get me the market caps" tasks end in a hand-typed spreadsheet that's stale the moment it's saved. This pipeline instead **discovers** the universe, **reverse-engineers** the auth needed to read it, and hands back a workbook that **recomputes itself** — so it survives a manual price edit or a change of dates.

- **Symbol discovery, not a hard-coded list.** Saudi REITs live in a contiguous ticker block (the `43xx` range). `discover.py` probes each code's chart endpoint and keeps the funds whose name contains "REIT" — the count (19, today) falls out of the data rather than being asserted.
- **Reverse-engineered Yahoo auth.** The useful `quoteSummary` endpoint (units outstanding, live market cap) is gated behind a "crumb" token that's only valid with the right cookies. The pipeline reproduces the exact handshake — a detail most tutorials get wrong.
- **Timezone-honest price matching.** Each target date is matched to the closest trading day *on or before* it using the **Asia/Riyadh** calendar, so a UTC midnight boundary never silently shifts a close by a day across a weekend or holiday.
- **A deliverable that stays live.** Only raw inputs (prices, units) are written as values; every derived figure — change, change %, market cap, sector totals — is an **Excel formula**. Edit a price and the whole sheet re-totals.
- **Trust, built in.** A second sheet documents the method and cross-checks several funds against Argaam, the Saudi Exchange and StockAnalysis; any missing datum is marked `n/a` rather than invented.

## The technical part, precisely

**The crumb/cookie handshake** (`extract.py`, reproduced standalone in `examples/yahoo_crumb_probe.py`):

```
1.  GET  fc.yahoo.com                         ─┐  collect session cookies
2.  GET  finance.yahoo.com/quote/<SYM>        ─┘
3.  GET  query2…/v1/test/getcrumb             →  crumb token (bound to those cookies)
4.  GET  query2…/v10/finance/quoteSummary/<SYM>?modules=price,defaultKeyStatistics&crumb=<token>
```

Without steps 1–3 the `quoteSummary` call returns `401/403`; with them it returns `sharesOutstanding` and the live market cap.

**Every number in the workbook, and how it's produced** — raw inputs vs. live formulas:

| Column | Source | How |
|---|---|---|
| Price (start) / Price (end) | value | Adjusted daily **close** from `/v8/finance/chart`, matched to the closest trading day ≤ target date (Asia/Riyadh) |
| Units outstanding | value | `sharesOutstanding.raw` from the crumb-authenticated `quoteSummary` |
| Change (SAR) | **formula** | `=F{r}-E{r}` |
| Change (%) | **formula** | `=(F{r}-E{r})/E{r}` |
| Market cap (end) | **formula** | `=I{r}*F{r}`  (units × end-date close) |
| Sector totals | **formula** | `=SUM(I{first}:I{last})`, `=SUM(J{first}:J{last})` |

Market cap for a *past* date is computed as `units × close`, deliberately in preference to Yahoo's "live" market cap (which reflects the day you happen to query). Rows are sorted largest-cap first; up/down changes are colour-banded green/red.

## How it works

```
        REIT_CODE_MIN..MAX  (default 4300..4360)
                   │
                   ▼
   discover.py ───────────────►  reit_list.json      [[code, name, last_price], …]
     probe /v8/finance/chart per code,
     keep names containing "REIT"
                   │
                   ▼
   extract.py ────────────────►  reit_data.json      {start_date, end_date, rows:[…]}
     crumb/cookie auth → getcrumb
     /v8/finance/chart      (daily closes, closest Riyadh trading day)
     /v10/…/quoteSummary    (units outstanding, live mcap)
                   │
                   ▼
   build_xlsx.py ─────────────►  KSA_REITs_price_and_market_cap.xlsx      (EN)
     one parameterised builder →  REITs_Arabie_Saoudite_prix_capitalisation.xlsx (FR)
     raw inputs as values; change / % / market cap / totals as live formulas
     sheet 2: methodology + cited sources
```

- **`discover.py`** — universe discovery by code-range scan (range overridable via `REIT_CODE_MIN` / `REIT_CODE_MAX`).
- **`extract.py`** — the crumb dance, windowed daily-close fetch, closest-trading-day matching, and per-fund market-cap / change derivation. Snapshot dates overridable via `REIT_START_DATE` / `REIT_END_DATE`.
- **`build_xlsx.py`** — a single, language-parameterised builder that emits both the English and French workbooks from one code path (labels in a `LABELS` table), keeping the two deliverables byte-for-byte consistent apart from language.
- **`examples/yahoo_crumb_probe.py`** — a minimal, standalone demo of the auth handshake for a single symbol.

## Tech stack

Python · httpx · openpyxl · Yahoo Finance public API · `zoneinfo` (Asia/Riyadh)

## Running it locally

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python3 discover.py      # -> reit_list.json  (find the REIT symbols)
python3 extract.py       # -> reit_data.json  (fetch prices + units)
python3 build_xlsx.py    # -> the two .xlsx workbooks (EN + FR)
```

Optional overrides (see `.env.example`): `REIT_CODE_MIN`, `REIT_CODE_MAX`, `REIT_START_DATE`, `REIT_END_DATE`. With no configuration it produces a point-in-time snapshot for the default dates.

## Data & privacy

This repository contains **code only**. It ships no scraped output — no `reit_list.json`, no `reit_data.json`, no `.xlsx`. Everything the workbook shows is **public Saudi stock-market data** pulled at run time from Yahoo Finance and cross-checked against the Saudi Exchange / Argaam / StockAnalysis. Run the three scripts and the deliverable is rebuilt locally.

## Project layout

```
discover.py      universe discovery (code-range scan)
extract.py       crumb auth + price/units fetch → reit_data.json
build_xlsx.py    bilingual, formula-driven workbook builder
examples/        standalone crumb-auth probe
```

## License

MIT — see [LICENSE](LICENSE).
