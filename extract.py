"""Step 2 — fetch closing prices + units outstanding for each discovered REIT.

Yahoo Finance gates its `quoteSummary` endpoint behind a "crumb" token that is
only issued once you hold the right cookies. This module reproduces that flow:

  1. GET fc.yahoo.com and a quote page  -> collect session cookies
  2. GET /v1/test/getcrumb             -> the crumb token (bound to those cookies)
  3. call /v10/finance/quoteSummary with ?crumb=<token>

For prices, the public /v8/finance/chart endpoint returns daily closes over a
window; each target date is matched to the closest trading day on or before it,
using the Asia/Riyadh calendar so a UTC boundary never shifts a bar by a day.

Reads:  reit_list.json (from discover.py)
Writes: reit_data.json

Env overrides (optional):
  REIT_START_DATE  first snapshot date, YYYY-MM-DD  (default 2026-04-01)
  REIT_END_DATE    second snapshot date, YYYY-MM-DD (default 2026-06-30)
"""
import datetime as dt
import json
import os
import time
from zoneinfo import ZoneInfo

import httpx

RIYADH = ZoneInfo("Asia/Riyadh")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

# Two snapshot dates to compare. Defaults make this a point-in-time example run.
START_DATE = dt.date.fromisoformat(os.getenv("REIT_START_DATE", "2026-04-01"))
END_DATE = dt.date.fromisoformat(os.getenv("REIT_END_DATE", "2026-06-30"))

# Fetch window: pad each side so we can walk back to the closest prior trading day.
p1 = int(dt.datetime.combine(START_DATE - dt.timedelta(days=12), dt.time(), RIYADH).timestamp())
p2 = int(dt.datetime.combine(END_DATE + dt.timedelta(days=3), dt.time(), RIYADH).timestamp())

c = httpx.Client(headers={"User-Agent": UA}, timeout=30, follow_redirects=True)

# --- crumb dance: collect cookies first, then fetch the crumb token ---
c.get("https://fc.yahoo.com")
c.get("https://finance.yahoo.com/quote/4330.SR")
CRUMB = c.get("https://query2.finance.yahoo.com/v1/test/getcrumb").text.strip()

reits = json.load(open("reit_list.json"))


def close_on_or_before(date_map, target, maxback=7):
    """Return (close, date) for `target`, walking back up to `maxback` days
    to skip weekends/holidays. Returns (None, None) if nothing is found."""
    d = target
    for _ in range(maxback + 1):
        if date_map.get(d) is not None:
            return date_map[d], d
        d = d - dt.timedelta(days=1)
    return None, None


rows = []
for code, name, _live in reits:
    sym = f"{code}.SR"
    row = {"code": code, "name": name, "symbol": sym}

    # --- daily closes over the window ---
    r = c.get(
        f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}",
        params={"period1": p1, "period2": p2, "interval": "1d"},
    )
    res = r.json()["chart"]["result"][0]
    ts = res["timestamp"]
    closes = res["indicators"]["quote"][0]["close"]
    date_map = {dt.datetime.fromtimestamp(t, RIYADH).date(): cl for t, cl in zip(ts, closes)}
    row["currency"] = res["meta"].get("currency")
    for key, tgt in (("start", START_DATE), ("end", END_DATE)):
        val, used = close_on_or_before(date_map, tgt)
        row[f"price_{key}"] = round(val, 4) if val else None
        row[f"date_{key}"] = used.isoformat() if used else None

    # --- units outstanding + live market cap (authenticated endpoint) ---
    time.sleep(0.2)
    q = c.get(
        f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{sym}",
        params={"modules": "price,defaultKeyStatistics", "crumb": CRUMB},
    )
    try:
        qr = q.json()["quoteSummary"]["result"][0]
        so = qr.get("defaultKeyStatistics", {}).get("sharesOutstanding", {}).get("raw")
        live_mcap = qr.get("price", {}).get("marketCap", {}).get("raw")
    except Exception:
        so, live_mcap = None, None
    row["shares_outstanding"] = so
    row["live_marketcap_yahoo"] = live_mcap

    # --- derived: market cap on the end date, and price evolution ---
    if so and row.get("price_end"):
        row["marketcap_end"] = round(so * row["price_end"])
    else:
        row["marketcap_end"] = None
    if row.get("price_start") and row.get("price_end"):
        row["pct_change"] = round((row["price_end"] - row["price_start"]) / row["price_start"] * 100, 2)
    else:
        row["pct_change"] = None

    rows.append(row)
    print(
        f"{code} {name[:34]:34s} start={row['price_start']} ({row['date_start']})  "
        f"end={row['price_end']} ({row['date_end']})  chg={row['pct_change']}%  "
        f"SO={so}  mcap_end={row['marketcap_end']}"
    )
    time.sleep(0.25)

payload = {
    "start_date": START_DATE.isoformat(),
    "end_date": END_DATE.isoformat(),
    "currency": "SAR",
    "rows": rows,
}
json.dump(payload, open("reit_data.json", "w"), ensure_ascii=False, indent=2)
print("\nSaved reit_data.json")
