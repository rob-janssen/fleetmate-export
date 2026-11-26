#!/usr/bin/env python3
"""
Fleetmate agenda -> Excel export.

Haalt de volledige agenda op via de Fleetmate/GO REST API (api/v2), bewaart een
ruwe CSV/JSON-export en vult het Excel-bestand opnieuw (oude data wordt eerst
leeggemaakt). Bedoeld om wekelijks via launchd te draaien.

Gebruik:
  python fleetmate_export.py                 # normale run (API -> CSV -> Excel)
  python fleetmate_export.py --discover      # zoek agenda-endpoints in de Swagger-spec
  python fleetmate_export.py --probe <pad>   # toon 3 records van een endpoint + veldnamen
  python fleetmate_export.py --from-json f   # vul Excel vanuit een eerder opgeslagen JSON
  python fleetmate_export.py --dry-run       # alles behalve het Excel-bestand wegschrijven
"""
import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from copy import copy

HERE = os.path.dirname(os.path.abspath(__file__))

MONTHS_NL = ["Januari", "Februari", "Maart", "April", "Mei", "Juni", "Juli",
             "Augustus", "September", "Oktober", "November", "December"]
TYPES = ["CFK", "APK", "Koeler", "Oplegger", "Onderhoud", "FRC", "Tachograaf", "Alarm"]
# Vaste kleur per type (hele rij in Overzicht). Staat in de code, dus raakt nooit kwijt als een type tijdelijk ontbreekt.
TYPE_COLORS = {"APK": "F7D6D6", "CFK": "D9E8F6", "Koeler": "D4EDE0", "Oplegger": "DEDCF4",
               "Tachograaf": "E5E3DA", "FRC": "F4D5E0", "Onderhoud": "F5E2C2", "Alarm": "EAD9EE"}

# Volgorde is belangrijk: eerste match wint.
TYPE_RULES = [
    ("CFK",        r"\bcfk\b"),
    ("APK",        r"\bapk\b"),
    ("FRC",        r"\bfrc\b"),
    ("Tachograaf", r"tacho"),
    ("Alarm",      r"\balarm"),
    ("Koeler",     r"koel"),
    ("Oplegger",   r"oplegger"),
    ("Onderhoud",  r"onderhoud|\bq[- ]?beurt|\bxg\b|\bxy\b"),
]

# Kandidaat-veldnamen als er in config.env niets is opgegeven.
FIELD_CANDIDATES = {
    "date": ["fromDate", "appointmentDate", "startDate", "start", "dateStart", "date", "plannedDate",
             "startTime", "beginDate", "datum", "from"],
    "license": ["vehicleLicense", "licensePlate", "license", "registration", "kenteken",
                "vehicle.license", "vehicle.licensePlate", "vehicle.registration"],
    "truck": ["vehicleNumber", "vehicleCode", "internalNumber", "fleetNumber", "vehicle.number",
              "vehicle.code", "vehicle.internalNumber", "truckNumber"],
    "description": ["title", "description", "subject", "omschrijving", "name", "remark", "notes"],
    "type": [],  # standaard afgeleid uit omschrijving
}

LOG_DIR = os.path.join(HERE, "logs")
EXPORT_DIR = os.path.join(HERE, "exports")
BACKUP_DIR = os.path.join(HERE, "backups")


# ---------------------------------------------------------------- utilities
def log(msg):
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    os.makedirs(LOG_DIR, exist_ok=True)
    with open(os.path.join(LOG_DIR, "export.log"), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def notify(title, msg):
    """macOS-notificatie; stil negeren als het niet lukt."""
    try:
        subprocess.run(["osascript", "-e",
                        f'display notification {json.dumps(msg)} with title {json.dumps(title)}'],
                       check=False, capture_output=True, timeout=10)
    except Exception:
        pass


def load_config(path):
    cfg = {}
    if not os.path.exists(path):
        sys.exit(f"Config niet gevonden: {path}")
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip().strip('"').strip("'")
    return cfg


def get_path(rec, dotted):
    cur = rec
    for part in dotted.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur


def pick(rec, configured, candidates):
    if configured:
        return get_path(rec, configured)
    for c in candidates:
        v = get_path(rec, c)
        if v not in (None, ""):
            return v
    return None


def parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):  # unix timestamp (s of ms)
        return dt.datetime.fromtimestamp(v / 1000 if v > 1e11 else v)
    s = str(v).strip().replace("Z", "")
    s = re.sub(r"[+-]\d{2}:?\d{2}$", "", s)  # tijdzone eraf
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d-%m-%Y %H:%M", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def classify(description, explicit=None):
    if explicit:
        for t in TYPES:
            if str(explicit).strip().lower() == t.lower():
                return t
    text = (description or "").lower()
    for t, rx in TYPE_RULES:
        if re.search(rx, text):
            return t
    return "Overig"


# ---------------------------------------------------------------- API
class Api:
    def __init__(self, base_url, api_key):
        self.base = base_url.rstrip("/") + "/"
        self.key = api_key

    def get(self, path, params=None, absolute=False):
        url = path if absolute else self.base + path.lstrip("/")
        if params:
            url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
        req = urllib.request.Request(url, headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": f"apikey {self.key}",
        })
        with urllib.request.urlopen(req, timeout=60) as r:
            body = r.read().decode("utf-8")
        return json.loads(body) if body else None

    def list_all(self, endpoint, where=None, page_size=500, max_pages=200):
        """Haalt alle records op. Werkt ook als het endpoint $limit/$offset negeert:
        dubbele ID's worden overgeslagen en we stoppen zodra een pagina niets nieuws oplevert."""
        out, seen, offset = [], set(), 0
        for _ in range(max_pages):
            params = {"$limit": page_size, "$offset": offset}
            if where:
                params["$where"] = where
            rows = unwrap(self.get(endpoint, params))
            new = 0
            for r in rows:
                rid = r.get("ID", r.get("id")) if isinstance(r, dict) else None
                key = rid if rid is not None else json.dumps(r, sort_keys=True, default=str)
                if key in seen:
                    continue
                seen.add(key)
                out.append(r)
                new += 1
            if new == 0 or len(rows) < page_size or len(rows) > page_size:
                break  # laatste pagina, of endpoint pagineert niet
            offset += page_size
        return out


def unwrap(data):
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for k in ("data", "items", "results", "records", "value", "rows"):
            if isinstance(data.get(k), list):
                return data[k]
    return [data] if data else []


# ---------------------------------------------------------------- discover / probe
AGENDA_RX = re.compile(r"agenda|appointment|planning|schedule|calendar|event|afspraak|workorder|workOrder|task", re.I)


def discover(api):
    root = re.sub(r"/api/v\d+/?$", "/", api.base)
    candidates = [root + p for p in (
        "swagger/swagger.json", "swagger/openapi.json", "swagger/v2/swagger.json", "swagger/doc.json",
        "swagger/api.json", "swagger/swagger.yaml", "api/v2/swagger.json", "api/v2/openapi.json",
        "api/swagger.json", "openapi.json")]
    spec = None
    for url in candidates:
        try:
            spec = api.get(url, absolute=True)
            if isinstance(spec, dict) and "paths" in spec:
                print(f"Spec gevonden: {url}\n")
                break
            spec = None
        except Exception:
            continue
    if not spec:
        print("Geen Swagger-spec gevonden op de standaardlocaties.")
        print(f"Open {root}swagger/ in je browser, zoek het agenda-endpoint en zet het in config.env (AGENDA_ENDPOINT).")
        print("Tip: in de browser devtools (Network) zie je welk .json-bestand de Swagger-pagina laadt.")
        return
    hits = [(p, m) for p, ops in spec["paths"].items() for m in ops if AGENDA_RX.search(p)]
    if not hits:
        print("Geen agenda-achtige paden gevonden. Alle GET-lijstpaden:")
        hits = [(p, m) for p, ops in spec["paths"].items() for m in ops if m.lower() == "get" and "{" not in p]
    for p, m in sorted(hits):
        summary = spec["paths"][p][m].get("summary", "") if isinstance(spec["paths"][p][m], dict) else ""
        print(f"  {m.upper():6} {p}   {summary}")
    print("\nZet het lijst-endpoint (zonder {id}) in config.env als AGENDA_ENDPOINT en draai --probe om de velden te zien.")


def probe(api, endpoint):
    rows = api.list_all(endpoint)
    print(json.dumps(rows[:3], indent=2, ensure_ascii=False, default=str))
    print(f"\nAantal records: {len(rows)}")
    if rows and isinstance(rows[0], dict):
        dkey = next((c for c in FIELD_CANDIDATES["date"] if get_path(rows[0], c) not in (None, "")), None)
        if dkey:
            dates = sorted(d for d in (parse_date(get_path(r, dkey)) for r in rows) if d)
            if dates:
                per_year = Counter(d.year for d in dates)
                print(f"Datumbereik ({dkey}): {dates[0]:%d-%m-%Y} t/m {dates[-1]:%d-%m-%Y}  "
                      f"per jaar: {dict(sorted(per_year.items()))}")
    if rows and isinstance(rows[0], dict):
        print("\nVelden:", ", ".join(rows[0].keys()))
        print("\nAutomatisch herkend:")
        for key, cands in FIELD_CANDIDATES.items():
            hit = next((c for c in cands if get_path(rows[0], c) not in (None, "")), None)
            print(f"  {key:12} -> {hit or '(niet gevonden, zet FIELD_' + key.upper() + ' in config.env)'}")


# ---------------------------------------------------------------- truck map
def load_truck_map(path):
    m = {}
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("kenteken") and row.get("truck"):
                    m[row["kenteken"].strip().upper()] = row["truck"].strip()
    return m


def save_truck_map(path, m):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["kenteken", "truck"])
        for k in sorted(m):
            w.writerow([k, m[k]])


def seed_truck_map_from_workbook(wb):
    """Leest bestaande kenteken -> Truck # koppelingen uit het huidige bestand (voor het leegmaken)."""
    votes = defaultdict(Counter)
    ws = wb["Overzicht 2026"]
    for r in ws.iter_rows(min_row=9, values_only=True):
        if r and len(r) >= 5 and r[3] and r[4]:
            votes[str(r[3]).strip().upper()][str(r[4]).strip()] += 1
    if "Per truck" in wb.sheetnames:
        for r in wb["Per truck"].iter_rows(min_row=4, values_only=True):
            if r and r[0] and r[1]:
                votes[str(r[0]).strip().upper()][str(r[1]).strip()] += 1
    return {k: c.most_common(1)[0][0] for k, c in votes.items()}


# ---------------------------------------------------------------- normalize
PLATE_RX = re.compile(r"kenteken\s+([A-Z0-9]{1,3}(?:-[A-Z0-9]{1,3}){2})", re.I)


def truthy(v, default=False):
    if v is None or str(v).strip() == "":
        return default
    return str(v).strip().lower() in ("1", "true", "ja", "yes", "y", "waar")


def int_set(v):
    return {int(x) for x in re.split(r"[,\s]+", v or "") if x.strip().lstrip("-").isdigit()}


def fetch_vehicle_map(api, endpoint):
    """kenteken -> vehicleNumber uit /vehicles (actieve voertuigen gaan voor)."""
    try:
        rows = unwrap(api.get(endpoint))
    except Exception as e:
        log(f"Let op: {endpoint} niet opgehaald ({type(e).__name__}: {e}); Truck # komt uit truck_map.csv")
        return {}
    m, rank = {}, {}
    for v in rows:
        if not isinstance(v, dict):
            continue
        lic = str(v.get("license") or "").strip().upper()
        num = str(v.get("vehicleNumber") or "").strip()
        if not lic or not num:
            continue
        r = (0 if v.get("isHistory") or v.get("deleted") else 1)
        if lic not in m or r > rank[lic]:
            m[lic], rank[lic] = num, r
    log(f"{len(m)} voertuigen met kenteken + nummer uit {endpoint}")
    return m


def normalize(records, cfg, truck_map, vehicle_map=None, now=None):
    """Filtert zoals de portal (toekomst, niet-historisch) en zet records om naar events.
    Elk event bewaart het bronrecord in '_rec' voor de CSV."""
    vehicle_map = vehicle_map or {}
    now = now or dt.datetime.now()
    fd = {k: cfg.get(f"FIELD_{k.upper()}", "") for k in FIELD_CANDIDATES}
    only_future = truthy(cfg.get("ONLY_FUTURE"), True)
    excl_hist = truthy(cfg.get("EXCLUDE_HISTORY"), True)
    events, skipped, stats = [], 0, Counter()
    for rec in records:
        if not isinstance(rec, dict):
            continue
        d = parse_date(pick(rec, fd["date"], FIELD_CANDIDATES["date"]))
        if d is None:
            skipped += 1
            continue
        if excl_hist and truthy(rec.get("isHistory")):
            stats["historie"] += 1
            continue
        if only_future and d <= now:
            stats["verleden"] += 1
            continue
        desc = pick(rec, fd["description"], FIELD_CANDIDATES["description"])
        desc = str(desc).strip() if desc else ""
        lic = pick(rec, fd["license"], FIELD_CANDIDATES["license"])
        lic = str(lic).strip().upper() if lic else ""
        if not lic:
            m = PLATE_RX.search(desc)
            lic = m.group(1).upper() if m else ""
        truck = pick(rec, fd["truck"], FIELD_CANDIDATES["truck"])
        truck = str(truck).strip() if truck else ""
        if not truck and lic:
            truck = vehicle_map.get(lic) or truck_map.get(lic, "")
        if truck and lic:
            truck_map[lic] = truck
        etype = classify(desc, pick(rec, fd["type"], []) if fd["type"] else None)
        if d.hour == 0 and d.minute == 0:
            d = dt.datetime(d.year, d.month, d.day)
        events.append({"Datum": d, "Dag": d.day, "Maand": MONTHS_NL[d.month - 1],
                       "Kenteken": lic, "Truck #": truck or None, "Type": etype, "Omschrijving": desc,
                       "itemType": rec.get("itemType"), "_rec": rec})
    events.sort(key=lambda e: (e["Datum"], e["Kenteken"], e["Type"]))
    if skipped:
        log(f"Let op: {skipped} records zonder herkenbare datum overgeslagen.")
    if stats:
        log("Weggefilterd: " + ", ".join(f"{k} {v}" for k, v in stats.items()))
    return events


def excel_selection(events, cfg):
    """Welke events in Excel komen (CSV bevat alles na het portal-filter)."""
    excl_types = int_set(cfg.get("EXCEL_EXCLUDE_ITEM_TYPES", "0,6"))
    year = cfg.get("YEAR", "").strip()
    until = cfg.get("EXCEL_UNTIL", "").strip()  # bv. 2027-01: alles t/m die maand
    uy, um = (int(x) for x in until.split("-")) if until else (None, None)
    sel = [e for e in events
           if (e["itemType"] is None or int(e["itemType"]) not in excl_types)
           and (not year or e["Datum"].year == int(year))
           and (not until or (e["Datum"].year, e["Datum"].month) <= (uy, um))]
    # Meerdere jaren? Dan krijgt Maand het jaartal erbij ("Januari 2027"), anders tellen jaren door elkaar.
    if len({e["Datum"].year for e in sel}) > 1:
        for e in sel:
            e["Maand"] = f"{MONTHS_NL[e['Datum'].month - 1]} {e['Datum'].year}"
    return sel


def year_label(events):
    ys = sorted({e["Datum"].year for e in events})
    if not ys:
        return str(dt.date.today().year)
    return str(ys[0]) if len(ys) == 1 else f"{ys[0]}–{ys[-1]}"


def retitle(ws, label):
    v = ws["A1"].value
    if isinstance(v, str) and re.search(r"\d{4}(–\d{4})?\s*$", v):
        ws["A1"] = re.sub(r"\d{4}(–\d{4})?\s*$", label, v)


# ---------------------------------------------------------------- excel
def style_of(cell):
    return {"font": copy(cell.font), "fill": copy(cell.fill), "border": copy(cell.border),
            "alignment": copy(cell.alignment), "number_format": cell.number_format,
            "protection": copy(cell.protection)}


def apply_style(cell, st):
    if not st:
        return
    cell.font, cell.fill, cell.border = st["font"], st["fill"], st["border"]
    cell.alignment, cell.number_format, cell.protection = st["alignment"], st["number_format"], st["protection"]


def clear_rows(ws, first_row):
    if ws.max_row >= first_row:
        ws.delete_rows(first_row, ws.max_row - first_row + 1)


def month_order(events):
    """Unieke Maand-labels in kalendervolgorde."""
    seen = {}
    for e in events:
        seen.setdefault(e["Maand"], (e["Datum"].year, e["Datum"].month))
    return sorted(seen, key=seen.get)


def write_overzicht(ws, events):
    HDR, FIRST = 8, 9
    row_style = [style_of(ws.cell(FIRST, c)) for c in range(1, 8)]
    type_style = {}
    for r in range(FIRST, ws.max_row + 1):
        v = ws.cell(r, 6).value
        if v and v not in type_style:
            type_style[v] = style_of(ws.cell(r, 6))
    clear_rows(ws, FIRST)
    cols = ["Datum", "Dag", "Maand", "Kenteken", "Truck #", "Type", "Omschrijving"]
    for i, e in enumerate(events):
        r = FIRST + i
        for c, key in enumerate(cols, start=1):
            cell = ws.cell(r, c, e[key])
            apply_style(cell, row_style[c - 1])
            if e["Type"] in TYPE_COLORS:
                from openpyxl.styles import PatternFill
                cell.fill = PatternFill("solid", fgColor="FF" + TYPE_COLORS[e["Type"]])
    last = FIRST + max(len(events), 1) - 1
    t = ws.tables["Events2026"]
    t.ref = f"A{HDR}:G{last}"
    if t.autoFilter is not None:
        t.autoFilter.ref = t.ref

    # Kerncijfers bovenaan
    order = month_order(events)
    months = Counter(e["Maand"] for e in events)
    days = Counter(e["Datum"].date() for e in events)
    ws["B3"] = len(events)
    ws["B4"] = len({e["Kenteken"] for e in events if e["Kenteken"]})
    if months:
        m, n = max(months.items(), key=lambda kv: (kv[1], -order.index(kv[0])))
        ws["B5"] = f"{m} ({n})"
    else:
        ws["B5"] = "-"
    if days:
        d, n = max(days.items(), key=lambda kv: (kv[1], -kv[0].toordinal()))
        ws["B6"] = f"{d.strftime('%d-%m-%Y')} ({n})"
    else:
        ws["B6"] = "-"


def write_per_maand(ws, events):
    from openpyxl.utils import get_column_letter
    FIRST = 4
    month_style = [style_of(ws.cell(FIRST, c)) for c in range(1, 11)]
    total_row = next((r for r in range(FIRST, ws.max_row + 1) if ws.cell(r, 1).value == "Totaal"), ws.max_row)
    total_style = [style_of(ws.cell(total_row, c)) for c in range(1, 11)]
    clear_rows(ws, FIRST)
    labels = month_order(events) or [MONTHS_NL[dt.date.today().month - 1]]
    src = "'Overzicht 2026'"
    for i, label in enumerate(labels):
        r = FIRST + i
        ws.cell(r, 1, label)
        ws.cell(r, 2, f"=COUNTIF({src}!C:C,A{r})")
        for j, t in enumerate(TYPES):
            ws.cell(r, 3 + j, f'=COUNTIFS({src}!C:C,A{r},{src}!F:F,"{t}")')
        for c in range(1, 11):
            apply_style(ws.cell(r, c), month_style[c - 1])
    last = FIRST + len(labels) - 1
    tr = last + 1
    ws.cell(tr, 1, "Totaal")
    for c in range(2, 11):
        col = get_column_letter(c)
        ws.cell(tr, c, f"=SUM({col}{FIRST}:{col}{last})")
    for c in range(1, 11):
        apply_style(ws.cell(tr, c), total_style[c - 1])
    if ws.column_dimensions["A"].width and ws.column_dimensions["A"].width < 16:
        ws.column_dimensions["A"].width = 16


def write_per_truck(ws, events, truck_map):
    HDR, FIRST = 3, 4
    row_style = [style_of(ws.cell(FIRST, c)) for c in range(1, 13)]
    clear_rows(ws, FIRST)
    plates = sorted({e["Kenteken"] for e in events if e["Kenteken"]})
    src = "'Overzicht 2026'"
    for i, p in enumerate(plates):
        r = FIRST + i
        ws.cell(r, 1, p)
        ws.cell(r, 2, truck_map.get(p))
        ws.cell(r, 3, f"=COUNTIF({src}!D:D,A{r})")
        for j, t in enumerate(TYPES):
            ws.cell(r, 4 + j, f'=COUNTIFS({src}!D:D,A{r},{src}!F:F,"{t}")')
        ws.cell(r, 12, f"=_xlfn.MINIFS({src}!A:A,{src}!D:D,A{r})")
        for c in range(1, 13):
            apply_style(ws.cell(r, c), row_style[c - 1])
    last = FIRST + max(len(plates), 1) - 1
    if "PerTruck" in ws.tables:
        t = ws.tables["PerTruck"]
        t.ref = f"A{HDR}:K{last}"
        if t.autoFilter is not None:
            t.autoFilter.ref = t.ref


def update_workbook(xlsx_path, events, truck_map, truck_map_path, dry_run=False):
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path)
    label = year_label(events)

    write_overzicht(wb["Overzicht 2026"], events)
    retitle(wb["Overzicht 2026"], label)
    if "Per maand" in wb.sheetnames:
        write_per_maand(wb["Per maand"], events)
        retitle(wb["Per maand"], label)
    if "Per truck" in wb.sheetnames:
        write_per_truck(wb["Per truck"], events, truck_map)
        retitle(wb["Per truck"], label)
    wb.calculation.fullCalcOnLoad = True  # Excel herberekent formules bij openen

    if dry_run:
        log(f"Dry-run: {len(events)} events voor Excel, bestand niet weggeschreven.")
        return

    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copy2(xlsx_path, os.path.join(BACKUP_DIR, f"{stamp}_{os.path.basename(xlsx_path)}"))
    backups = sorted(f for f in os.listdir(BACKUP_DIR) if f.endswith(".xlsx"))
    for old in backups[:-8]:  # laatste 8 back-ups bewaren
        os.remove(os.path.join(BACKUP_DIR, old))

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(xlsx_path))
    os.close(fd)
    wb.save(tmp)
    os.replace(tmp, xlsx_path)
    save_truck_map(truck_map_path, truck_map)
    log(f"Excel bijgewerkt: {xlsx_path} ({len(events)} events)")


CSV_COLS = ["ID", "fromName", "toName", "fromDate", "toDate", "title", "itemType", "itemId",
            "vehicleId", "license", "vehicleNumber", "scheduleId", "hasSchedule"]


def write_raw_exports(records, events):
    """JSON = alles uit de API; CSV = portal-formaat na filter, aangevuld met kenteken/Truck #/Type."""
    os.makedirs(EXPORT_DIR, exist_ok=True)
    day = dt.date.today().isoformat()
    with open(os.path.join(EXPORT_DIR, f"agenda_raw_{day}.json"), "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=1, default=str)
    path = os.path.join(EXPORT_DIR, f"agenda_{day}.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(CSV_COLS + ["Type"])
        for e in events:
            rec = e["_rec"]
            row = []
            for c in CSV_COLS:
                v = rec.get(c, "")
                if c in ("fromDate", "toDate"):
                    pd = parse_date(v)
                    v = pd.strftime("%d/%m/%Y %H:%M:%S") if pd else v
                elif c == "license":
                    v = v or e["Kenteken"]
                elif c == "vehicleNumber":
                    v = v or (e["Truck #"] or "")
                row.append("" if v is None else v)
            w.writerow(row + [e["Type"]])
    for old in sorted(os.listdir(EXPORT_DIR))[:-24]:  # ca. 12 weken bewaren
        os.remove(os.path.join(EXPORT_DIR, old))
    return path


# ---------------------------------------------------------------- pipeline (Azure Function)
def run_pipeline(cfg, xlsx, truck_map_path, dry_run=False):
    """Zelfde stappen als main(), maar zonder CLI en met een config-dict. Geeft een samenvatting terug."""
    key = cfg.get("FLEETMATE_API_KEY", "")
    if not key or key.startswith("<"):
        raise RuntimeError("FLEETMATE_API_KEY ontbreekt in de app-instellingen.")
    api = Api(cfg.get("FLEETMATE_BASE_URL") or "https://srv.fleetmate.online/api/v2/", key)
    endpoint = cfg.get("AGENDA_ENDPOINT") or "agendas"
    log(f"Ophalen: {endpoint}")
    if truthy(cfg.get("PAGINATE"), False):
        records = api.list_all(endpoint, where=cfg.get("AGENDA_WHERE") or None,
                               page_size=int(cfg.get("PAGE_SIZE") or "500"))
    else:
        records = unwrap(api.get(endpoint))
    log(f"{len(records)} records opgehaald")
    vehicle_map = {}
    if cfg.get("VEHICLES_ENDPOINT", "vehicles"):
        vehicle_map = fetch_vehicle_map(api, cfg.get("VEHICLES_ENDPOINT") or "vehicles")

    truck_map = load_truck_map(truck_map_path)
    if os.path.exists(xlsx):
        from openpyxl import load_workbook
        truck_map = {**seed_truck_map_from_workbook(load_workbook(xlsx)), **truck_map}
    truck_map.update(vehicle_map)

    events = normalize(records, cfg, truck_map, vehicle_map)
    csv_path = write_raw_exports(records, events)
    log(f"CSV geschreven: {csv_path} ({len(events)} rijen)")
    sel = excel_selection(events, cfg)
    unknown = Counter((e["Omschrijving"].split(" voor ")[0][:40] or "(lege titel)")
                      for e in sel if e["Type"] == "Overig")
    update_workbook(xlsx, sel, truck_map, truck_map_path, dry_run=dry_run)
    return {"events": len(sel), "csv": csv_path, "months": month_order(sel),
            "unknown": dict(unknown), "records": len(records)}


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Fleetmate agenda -> Excel")
    ap.add_argument("--config", default=os.path.join(HERE, "config.env"))
    ap.add_argument("--discover", action="store_true")
    ap.add_argument("--probe", metavar="ENDPOINT")
    ap.add_argument("--from-json", metavar="FILE")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    cfg = load_config(args.config)
    xlsx = cfg.get("XLSX_PATH") or "werkplaats_agenda_2026.xlsx"
    xlsx = xlsx if os.path.isabs(xlsx) else os.path.join(HERE, os.path.expanduser(xlsx))
    truck_map_path = os.path.join(HERE, "truck_map.csv")
    vehicle_map = {}

    if args.from_json:
        with open(args.from_json, encoding="utf-8") as f:
            records = unwrap(json.load(f))
    else:
        key = cfg.get("FLEETMATE_API_KEY", "")
        if not key or key.startswith("<"):
            sys.exit("Vul eerst FLEETMATE_API_KEY in config.env in.")
        api = Api(cfg.get("FLEETMATE_BASE_URL", "https://srv.fleetmate.online/api/v2/"), key)
        if args.discover:
            return discover(api)
        if args.probe:
            return probe(api, args.probe)
        endpoint = cfg.get("AGENDA_ENDPOINT", "") or "agendas"
        log(f"Ophalen: {endpoint}")
        if truthy(cfg.get("PAGINATE"), False):
            records = api.list_all(endpoint, where=cfg.get("AGENDA_WHERE") or None,
                                   page_size=int(cfg.get("PAGE_SIZE", "500")))
        else:
            records = unwrap(api.get(endpoint))  # /agendas geeft alles in één keer terug
        log(f"{len(records)} records opgehaald")
        if cfg.get("VEHICLES_ENDPOINT", "vehicles"):
            vehicle_map = fetch_vehicle_map(api, cfg.get("VEHICLES_ENDPOINT", "vehicles"))

    truck_map = load_truck_map(truck_map_path)
    if os.path.exists(xlsx):  # bestaande koppelingen kenteken -> Truck # uit Excel meenemen
        from openpyxl import load_workbook
        truck_map = {**seed_truck_map_from_workbook(load_workbook(xlsx)), **truck_map}
    truck_map.update(vehicle_map)  # Fleetmate is leidend

    events = normalize(records, cfg, truck_map, vehicle_map)
    csv_path = write_raw_exports(records, events)
    log(f"CSV geschreven: {csv_path} ({len(events)} rijen, vergelijk met portal-export)")

    sel = excel_selection(events, cfg)
    unknown = Counter((e["Omschrijving"].split(" voor ")[0][:40] or "(lege titel)") for e in sel if e["Type"] == "Overig")
    if unknown:
        log(f"Let op: {sum(unknown.values())} events met Type 'Overig': " +
            ", ".join(f"{k} ({v})" for k, v in unknown.most_common(5)))
    update_workbook(xlsx, sel, truck_map, truck_map_path, dry_run=args.dry_run)


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")[:500]
        log(f"FOUT: HTTP {e.code} {e.reason} - {body}")
        notify("Fleetmate export mislukt", f"HTTP {e.code} {e.reason}")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as e:
        log(f"FOUT: {type(e).__name__}: {e}")
        notify("Fleetmate export mislukt", f"{type(e).__name__}: {e}")
        sys.exit(1)