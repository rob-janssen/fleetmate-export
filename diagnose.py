#!/usr/bin/env python3
"""Eenmalige diagnose: welke agenda-endpoints/parameters geven hoeveel records terug?
Gebruik: .venv/bin/python diagnose.py   (in dezelfde map als config.env)"""
import json
import os
import urllib.error
from collections import Counter

import fleetmate_export as fe

HERE = os.path.dirname(os.path.abspath(__file__))
cfg = fe.load_config(os.path.join(HERE, "config.env"))
api = fe.Api(cfg.get("FLEETMATE_BASE_URL", "https://srv.fleetmate.online/api/v2/"), cfg["FLEETMATE_API_KEY"])

FROM, TO = "2020-01-01", "2030-12-31"
attempts = [
    ("agendas/list", None),
    ("agendas/list", {"$limit": 5000}),
    ("agendas/list", {"$where": f"fromDate gt '{FROM}'"}),
    ("agendas", None),
    ("agendas/schedule", None),
    ("agendas/widget", None),
]
for a, b in [("from", "to"), ("fromDate", "toDate"), ("start", "end"), ("startDate", "endDate"), ("dateFrom", "dateTo")]:
    attempts.append(("agendas/schedule", {a: FROM, b: TO}))
attempts.append(("agendas/list", {"from": FROM, "to": TO}))
attempts.append(("agendas/list", {"fromDate": FROM, "toDate": TO}))


def summarize(rows):
    if not rows:
        return "0 records"
    first = rows[0] if isinstance(rows[0], dict) else {}
    ids = [r.get("ID", r.get("id")) for r in rows if isinstance(r, dict)]
    dkey = next((k for k in ("fromDate", "start", "startDate", "date", "appointmentDate") if k in first), None)
    out = [f"{len(rows)} records, {len(set(ids))} unieke ID's"]
    if dkey:
        ds = sorted(d for d in (fe.parse_date(r.get(dkey)) for r in rows if isinstance(r, dict)) if d)
        if ds:
            out.append(f"{dkey}: {ds[0]:%d-%m-%Y} t/m {ds[-1]:%d-%m-%Y}, per jaar {dict(sorted(Counter(d.year for d in ds).items()))}")
    if "itemType" in first:
        out.append(f"itemType: {dict(sorted(Counter(r.get('itemType') for r in rows).items()))}")
    if "hasSchedule" in first:
        out.append(f"hasSchedule=true: {sum(1 for r in rows if r.get('hasSchedule'))}")
    out.append(f"velden: {', '.join(first.keys())}")
    return "\n      ".join(out)


for ep, params in attempts:
    label = ep + ("?" + "&".join(f"{k}={v}" for k, v in params.items()) if params else "")
    try:
        data = api.get(ep, params)
        rows = fe.unwrap(data)
        extra = ""
        if isinstance(data, dict) and not any(isinstance(data.get(k), list) for k in ("data", "items", "results", "records", "value", "rows")):
            extra = f"\n      (object-response, keys: {', '.join(list(data.keys())[:15])})"
        print(f"OK   {label}\n      {summarize(rows)}{extra}")
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code} {label}\n      {e.read().decode('utf-8', 'replace')[:200]}")
    except Exception as e:
        print(f"FOUT {label}\n      {type(e).__name__}: {e}")
