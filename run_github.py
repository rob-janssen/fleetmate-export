"""Start de export vanuit GitHub Actions. Sjabloon staat in data/, resultaat komt in out/ (wordt als artifact bewaard)."""
import datetime as dt
import os
import shutil
from zoneinfo import ZoneInfo

import graph_mail
import job


class FileStore:
    """Zelfde interface als BlobStore, maar op de map data/ (invoer) en out/ (uitvoer)."""
    def download(self, name, path):
        src = os.path.join(os.environ.get("DATA_DIR", "data"), name)
        if not os.path.exists(src):
            return False
        shutil.copy(src, path)
        return True

    def upload(self, path, name):
        dst = os.path.join("out", name)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy(path, dst)


def main():
    manual = os.environ.get("GITHUB_EVENT_NAME") == "workflow_dispatch"
    if not manual:
        # Cron staat op 05:00 en 06:00 UTC; alleen de run op SEND_HOUR (Amsterdamse tijd) doet iets.
        send_hour = int(os.environ.get("SEND_HOUR", "7"))
        now = dt.datetime.now(ZoneInfo("Europe/Amsterdam"))
        if now.hour != send_hour:
            print(f"Overgeslagen: lokaal uur is {now.hour}, SEND_HOUR is {send_hour}.")
            return
    everyone = (not manual) or os.environ.get("NAAR") == "iedereen"
    to = graph_mail.split_addresses(os.environ.get("MAIL_TO" if everyone else "MAIL_ADMIN", ""))
    job.run_job_safe(to, store=FileStore())


if __name__ == "__main__":
    main()
