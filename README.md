# Fleetmate agenda → Excel (wekelijks, donderdag 07:00)

Haalt de volledige agenda op uit de Fleetmate API, slaat een CSV op in `exports/`,
maakt `werkplaats_agenda_2026.xlsx` leeg en vult hem opnieuw:
- **Overzicht 2026**: alle events met Datum, Dag, Maand, Kenteken, Truck #, Type en Omschrijving, plus de kerncijfers bovenaan
- **Per maand** en **Per truck**: rijen worden opnieuw opgebouwd (nieuwe maanden en kentekens komen er automatisch bij), met dezelfde formules
- **Type** wordt afgeleid uit de omschrijving (CFK, APK, Koeler, Oplegger, Onderhoud, FRC, Tachograaf, Alarm). Wat niet herkend wordt, krijgt "Overig"
- **Truck #** komt uit de API (als dat veld bestaat) of uit `truck_map.csv`. Die tabel wordt bij de eerste run gevuld vanuit je huidige Excel en kun je zelf aanvullen

## Installatie (eenmalig)

Zet de map buiten Desktop/Documents/Downloads (bv. `~/dev/fleetmate-export`), anders blokkeert macOS launchd.

```bash
cd ~/dev/fleetmate-export
nano config.env                         # API-key invullen
./install.sh                            # venv + launchd-agent
```

## Agenda-endpoint vinden

```bash
.venv/bin/python fleetmate_export.py --discover            # zoekt agenda-endpoints in de Swagger-spec
.venv/bin/python fleetmate_export.py --probe appointments   # toont 3 records + herkende velden
```

Zet het juiste endpoint in `config.env` bij `AGENDA_ENDPOINT`. Worden velden niet automatisch
herkend, vul dan `FIELD_DATE`, `FIELD_LICENSE`, `FIELD_TRUCK` en `FIELD_DESCRIPTION` in.

## Testen

```bash
.venv/bin/python fleetmate_export.py --dry-run     # ophalen + CSV, Excel ongemoeid
.venv/bin/python fleetmate_export.py               # volledige run
launchctl kickstart -k gui/$(id -u)/nl.rwmconsultancy.fleetmate-export   # via launchd
```

## Goed om te weten
- Je Mac moet aan staan. Sliep hij om 07:00, dan draait de export bij het ontwaken. Stond hij uit, dan wordt de run overgeslagen.
- Voor elke run gaat er een back-up naar `backups/` (de laatste 8 blijven bewaard).
- Log: `logs/export.log`. Bij een fout krijg je een macOS-notificatie.
- Excel herberekent de formules bij het openen.
