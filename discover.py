"""Step 1 — discover the Saudi REITs listed on Tadawul.

Saudi REIT funds occupy a contiguous ticker block on the Saudi Exchange
(the 43xx range). We probe each candidate symbol's public Yahoo Finance chart
endpoint and keep the funds whose long/short name contains "REIT".

Output: reit_list.json  ->  [[code, name, last_price], ...]

Env overrides (optional):
  REIT_CODE_MIN  first ticker code to probe   (default 4300)
  REIT_CODE_MAX  last  ticker code to probe   (default 4360, exclusive)
"""
import json
import os
import time

import httpx

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CODE_MIN = int(os.getenv("REIT_CODE_MIN", "4300"))
CODE_MAX = int(os.getenv("REIT_CODE_MAX", "4360"))

client = httpx.Client(headers={"User-Agent": UA}, timeout=20)

found = []
for code in range(CODE_MIN, CODE_MAX):
    sym = f"{code}.SR"
    try:
        r = client.get(
            f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}",
            params={"range": "1d", "interval": "1d"},
        )
        if r.status_code != 200:
            continue
        meta = r.json()["chart"]["result"][0]["meta"]
        name = meta.get("longName") or meta.get("shortName") or ""
        if "REIT" in name.upper():
            found.append((code, name, meta.get("regularMarketPrice")))
            print(f"{code}\t{name}\t{meta.get('regularMarketPrice')}")
    except Exception:
        # A non-REIT / delisted code returns junk or errors — skip it quietly.
        pass
    time.sleep(0.15)

print(f"\nTOTAL REITs found: {len(found)}")
json.dump(found, open("reit_list.json", "w"), ensure_ascii=False, indent=2)
print("Saved reit_list.json")
