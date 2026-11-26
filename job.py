"""De eigenlijke export-taak: sjabloon uit Blob halen, bijwerken, terugzetten en mailen."""
import datetime as dt
import os
import tempfile
import time
import traceback

import fleetmate_export as core
import graph_mail

TZ = "Europe/Amsterdam"
CONTAINER = os.environ.get("EXPORT_CONTAINER", "fleetmate-export")
XLSX_BLOB = "werkplaats_agenda_2026.xlsx"
MAP_BLOB = "truck_map.csv"

CONFIG_KEYS = ["FLEETMATE_BASE_URL", "FLEETMATE_API_KEY", "AGENDA_ENDPOINT", "VEHICLES_ENDPOINT",
               "ONLY_FUTURE", "EXCLUDE_HISTORY", "EXCEL_EXCLUDE_ITEM_TYPES", "PAGINATE", "AGENDA_WHERE",
               "YEAR", "EXCEL_UNTIL", "PAGE_SIZE", "FIELD_DATE", "FIELD_LICENSE", "FIELD_TRUCK",
               "FIELD_DESCRIPTION", "FIELD_TYPE"]


class BlobStore:
    def __init__(self):
        from azure.storage.blob import BlobServiceClient
        svc = BlobServiceClient.from_connection_string(
            os.environ.get("EXPORT_STORAGE_CONNECTION") or os.environ["AzureWebJobsStorage"])
        self.c = svc.get_container_client(CONTAINER)
        if not self.c.exists():
            self.c.create_container()

    def download(self, name, path):
        b = self.c.get_blob_client(name)
        if not b.exists():
            return False
        with open(path, "wb") as f:
            f.write(b.download_blob().readall())
        return True

    def upload(self, path, name):
        with open(path, "rb") as f:
            self.c.get_blob_client(name).upload_blob(f, overwrite=True)


def build_cfg():
    return {k: os.environ.get(k, "").strip() for k in CONFIG_KEYS}


def run_job(to, store=None, send=graph_mail.send):
    os.environ["TZ"] = TZ
    if hasattr(time, "tzset"):
        time.tzset()
    store = store or BlobStore()
    work = tempfile.mkdtemp(prefix="fleetmate-")
    core.LOG_DIR = os.path.join(work, "logs")
    core.EXPORT_DIR = os.path.join(work, "exports")
    core.BACKUP_DIR = os.path.join(work, "backups")
    xlsx = os.path.join(work, XLSX_BLOB)
    tmap = os.path.join(work, MAP_BLOB)
    if not store.download(XLSX_BLOB, xlsx):
        raise RuntimeError(f"Sjabloon '{XLSX_BLOB}' ontbreekt in container '{CONTAINER}'. Upload eerst de Excel.")
    store.download(MAP_BLOB, tmap)

    result = core.run_pipeline(build_cfg(), xlsx, tmap)

    day = dt.date.today().isoformat()
    store.upload(xlsx, XLSX_BLOB)
    if os.path.exists(tmap):
        store.upload(tmap, MAP_BLOB)
    store.upload(xlsx, f"archief/{day}_werkplaats_agenda.xlsx")
    store.upload(result["csv"], f"archief/{day}_agenda.csv")

    with open(xlsx, "rb") as f:
        content = f.read()
    body = (f"Hierbij de werkplaats agenda van {dt.date.today().strftime('%d-%m-%Y')}.\n\n"
            f"Events: {result['events']}\nPeriode: {', '.join(result['months']) or '-'}\n")
    if result["unknown"]:
        body += "\nLet op, niet-herkende types (staan als 'Overig'): " + \
                ", ".join(f"{k} ({v})" for k, v in result["unknown"].items()) + "\n"
    send(to, f"Werkplaats agenda {day}", body, [(f"werkplaats_agenda_{day}.xlsx", content)])
    return result


def run_job_safe(to, store=None):
    """Bij een fout: mail naar de beheerder en de fout opnieuw gooien zodat Azure de run als mislukt toont."""
    try:
        return run_job(to, store=store)
    except Exception as e:
        admin = graph_mail.split_addresses(os.environ.get("MAIL_ADMIN", ""))
        if admin:
            try:
                graph_mail.send(admin, "Fleetmate export MISLUKT",
                                f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-1500:]}")
            except Exception:
                pass
        raise
