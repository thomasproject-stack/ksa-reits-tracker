"""Minimal standalone demo of Yahoo Finance's crumb/cookie authentication.

Run this on its own to see the three-step handshake that `extract.py` relies on:
collect cookies, fetch the crumb token bound to them, then call the otherwise
403-gated quoteSummary endpoint. Handy for probing a single symbol.

    python examples/yahoo_crumb_probe.py
"""
import httpx

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
c = httpx.Client(headers={"User-Agent": UA}, timeout=20, follow_redirects=True)

# 1) hit a page to receive session cookies
c.get("https://fc.yahoo.com")
c.get("https://finance.yahoo.com/quote/4330.SR")

# 2) fetch the crumb token (only valid alongside the cookies above)
r = c.get("https://query2.finance.yahoo.com/v1/test/getcrumb")
print("CRUMB:", repr(r.text[:60]), "status", r.status_code)

# 3) call the authenticated quoteSummary endpoint
crumb = r.text.strip()
q = c.get(
    "https://query2.finance.yahoo.com/v10/finance/quoteSummary/4330.SR",
    params={"modules": "price,defaultKeyStatistics", "crumb": crumb},
)
print("quoteSummary status", q.status_code)
print(q.text[:600])
