# Fleetmate agenda → Excel, wekelijks per mail (GitHub Actions)

Elke donderdag 07:00 (Amsterdamse tijd) haalt GitHub de agenda uit Fleetmate, werkt `werkplaats_agenda_2026.xlsx`
bij (kleuren per type, tot en met `EXCEL_UNTIL`) en mailt hem via Microsoft Graph. Excel + CSV van elke run staan
90 dagen als artifact bij de run.

## Eenmalig instellen
1. **Repo:** privé repository, deze hele map erin zetten (de Excel bevat bedrijfsgegevens). `config.env` hoort er NIET in.
2. **Mail-app (Entra ID):**
   ```bash
   az ad app create --display-name fleetmate-export-mail --query appId -o tsv          # CLIENT_ID
   az ad sp create --id <CLIENT_ID>
   az ad app permission add --id <CLIENT_ID> --api 00000003-0000-0000-c000-000000000000 \
      --api-permissions b633e1c5-b582-4048-a93e-9f11b44c7e96=Role
   az ad app permission admin-consent --id <CLIENT_ID>
   az ad app credential reset --id <CLIENT_ID> --years 2 --query password -o tsv       # CLIENT_SECRET
   az account show --query tenantId -o tsv                                              # TENANT_ID
   ```
   Beperk de app tot één mailbox met `New-ApplicationAccessPolicy ... -AccessRight RestrictAccess` (Exchange Online).
3. **Secrets:** repo > Settings > Secrets and variables > Actions > New repository secret:
   `FLEETMATE_API_KEY` (vernieuw de oude key!), `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`,
   `MAIL_FROM` (bv. agenda@rwmconsultancy.nl), `MAIL_TO` (adressen gescheiden door `;`), `MAIL_ADMIN` (jij, voor foutmails).
4. **Testen:** tab Actions > Fleetmate export > Run workflow > `admin` (mail alleen naar jou) of `iedereen`.
5. **Mac-versie uitzetten** (anders dubbele runs): `launchctl bootout gui/$(id -u)/nl.rwmconsultancy.fleetmate-export`

## Aanpassen
- Periode: `EXCEL_UNTIL` in `.github/workflows/fleetmate-export.yml` (nu 2027-01).
- Kleuren: `TYPE_COLORS` bovenin `fleetmate_export.py`.
- Truck-koppelingen aanvullen: `truck_map.csv`.

## Goed om te weten
- Fleetmate moet de API vanaf GitHub-IP's toestaan (geen vast IP).
- GitHub start geplande runs soms te laat; handmatig starten kan altijd.
- Lokaal draaien kan nog: `config.env.example` kopiëren naar `config.env`, key invullen, `python fleetmate_export.py --dry-run`.
