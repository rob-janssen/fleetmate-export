#!/usr/bin/env python3
"""Diagnose ronde 2: /agendas uitgesplitst, /agendas/schedule en /vehicles.
Gebruik: .venv/bin/python diagnose2.py"""
import datetime as dt
import json
import os
import urllib.error
from collections import Counter, defaultdict

import fleetmate_export as fe

HERE = os.path.dirname(os.path.abspath(__file__))
cfg = fe.load_config(os.path.join(HERE, "config.env"))
api = fe.Api(cfg.get("FLEETMATE_BASE_URL", "https://srv.fleetmate.online/api/v2/"), cfg["FLEETMATE_API_KEY"])
today = dt.datetime.combine(dt.date.today(), dt.time())


def fetch(ep):
    try:
        return fe.unwrap(api.get(ep))
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code} {ep}: {e.read().decode('utf-8', 'replace')[:200]}")
    except Exception as e:
        print(f"FOUT {ep}: {type(e).__name__}: {e}")
    return []


def short(v, n=70):
    s = " ".join(str(v or "").split())
    return s if len(s) <= n else s[:n] + "…"


# ---- /agendas
rows = fetch("agendas")
print(f"=== agendas: {len(rows)} records")
if rows:
    def d(r): return fe.parse_date(r.get("fromDate"))
    fut = [r for r in rows if d(r) and d(r) >= today]
    nohist = [r for r in rows if not r.get("isHistory")]
    both = [r for r in fut if not r.get("isHistory")]
    print(f"  vanaf vandaag: {len(fut)} | isHistory=false: {len(nohist)} | beide: {len(both)}")
    print(f"  isHistory waarden: {dict(Counter(str(r.get('isHistory')) for r in rows))}")
    print(f"  branchId: {dict(Counter(str(r.get('branchId')) for r in rows))}")
    print(f"  fromName: {dict(Counter(str(r.get('fromName')) for r in rows).most_common(8))}")
    by = defaultdict(list)
    for r in rows:
        by[r.get("itemType")].append(r)
    for t in sorted(by, key=lambda x: (x is None, str(x))):
        rs = by[t]
        f = sum(1 for r in rs if d(r) and d(r) >= today)
        h = sum(1 for r in rs if r.get("isHistory"))
        print(f"\n  itemType {t}: {len(rs)} (vanaf vandaag {f}, isHistory {h})")
        for r in rs[:3]:
            print(f"     - {r.get('fromDate', '')[:10]} | title: {short(r.get('title'))} | descr: {short(r.get('description'), 50)} | itemId {r.get('itemId')}")

# ---- /agendas/schedule
rows = fetch("agendas/schedule")
print(f"\n=== agendas/schedule: {len(rows)} records")
if rows:
    print(f"  velden: {', '.join(rows[0].keys()) if isinstance(rows[0], dict) else type(rows[0])}")
    for r in rows[:3]:
        print("   ", short(json.dumps(r, ensure_ascii=False, default=str), 300))

# ---- /vehicles (voor Truck #)
for ep in ("vehicles", "vehicles/list"):
    rows = fetch(ep)
    if rows:
        print(f"\n=== {ep}: {len(rows)} records")
        print(f"  velden: {', '.join(rows[0].keys())}")
        print("   ", short(json.dumps(rows[0], ensure_ascii=False, default=str), 300))
        break
